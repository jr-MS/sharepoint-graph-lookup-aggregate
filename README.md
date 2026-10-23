# SharePoint lookup aggregation

One Azure Function replaces the Microsoft Graph list call that returns `null` for person and lookup columns.

This playbook is for someone who has not used Azure Functions before. It covers two things:

1. How the function wraps Graph.
2. How to call `POST /api/aggregate`.

The sample call is the JSON body in [The API](#the-api). Send it with curl. A Bruno collection in this repo is an optional client for the same body.

Do these six steps in order. Each one is written out below.

1. Sign in to the Azure subscription where this app should live.
2. Create a Linux Function App on Basic B1 and turn on its system-assigned identity. Move to Premium EP1 only if B1 is too slow or times out.
3. Grant that identity `Sites.Selected`, then grant it read on this one list's site.
4. Deploy this repository. The portal wizard does not contain this code.
5. Call `mode=mock` and `top=2` to prove the URL and key. This does not touch SharePoint.
6. Call `mode=graph` with `top=5`, then `top=1500`.

You need permission to create a Function App, and an admin who can grant Microsoft Graph application roles and SharePoint site permission. The function key is created by Azure after deploy. It is not in this repository.

## The problem

This list call is what fails today:

```http
GET https://graph.microsoft.com/v1.0/sites/{hostname}.sharepoint.com,{site-collection-id},{site-id}/lists/{list-id}/items?$expand=fields($select=Title,OwnerLookupId,Owner,CategoryLookupId,Category)
```

SharePoint's list-view lookup threshold is 12 columns. `Author` and `Editor` count. Past that cutoff, a collection `GET /items` returns HTTP 200 and sets both projections of the column to `null`:

```json
"OwnerLookupId": null,
"Owner": null
```

`LookupId` is not a second stored field. It is the id projection of the same person or lookup column. The display name is the other projection. The collection call drops both together.

The same item, read alone, returns both:

```http
GET /sites/{site-id}/lists/{list-id}/items/{item-id}?$expand=fields($select=OwnerLookupId,Owner)
```

```json
"OwnerLookupId": "12",
"Owner": "Example User"
```

Selecting fewer columns does not lift the limit. The threshold is the number of lookup-type columns on the list, not the length of `$select`.

## What you deploy

A Function App is an Azure website that runs one small HTTP API. This repo is that API.

| Piece | What it is |
| --- | --- |
| Function App | The host. Its address looks like `https://YOUR-FUNCTION-APP.azurewebsites.net`. |
| Function | One route inside that host. This repo has one: `aggregate`. |
| Function key | The password for that route. It is not a Graph token. |
| App setting | A named value stored on the Function App, such as `GRAPH_SITE_ID`. |

The caller sends one POST. The function then:

1. Reads item ids with `GET /items?$select=id`, 999 ids per page.
2. Splits those ids into groups of 20. Graph rejects a `$batch` with more than 20 requests.
3. Sends up to 3 groups at a time to `POST https://graph.microsoft.com/v1.0/$batch`.
4. Each inner request is `GET /items/{id}?$expand=fields($select=...)`.
5. Returns one JSON document with every selected field, including `LookupId` and the display name.

For 1,500 items that is 75 batches. A test of this same pattern returned in about 54 seconds. That was not a call against this production list. Check `result.failures` before you trust `result.items`.

## The API

```http
POST https://YOUR-FUNCTION-APP.azurewebsites.net/api/aggregate?code=YOUR-FUNCTION-KEY
Content-Type: application/json
```

The same key can be sent as a header instead of a query parameter. Prefer the header. Query strings are copied into logs.

```http
x-functions-key: YOUR-FUNCTION-KEY
```

Request body. Replace `fields` with the internal names on your list. Include both the display name and the `LookupId` when you need both.

```json
{
  "mode": "graph",
  "top": 1500,
  "fields": [
    "Title",
    "Owner",
    "OwnerLookupId",
    "Category",
    "CategoryLookupId"
  ]
}
```

| Field | Required | Meaning |
| --- | --- | --- |
| `mode` | no | `graph` reads SharePoint. `mock` returns synthetic rows and does not call Graph. Default is `graph`. |
| `top` | no | Maximum items to read. Default is 1500 in `graph` mode. Maximum is 2000. |
| `fields` | no | Column internal names. If omitted, the function uses the list above. Replace those names with the columns on your list. `OwnerLookupId` and `Owner` are both required if you need both projections of that column. |
| `siteId` | no | Overrides app setting `GRAPH_SITE_ID`. |
| `listId` | no | Overrides app setting `GRAPH_LIST_ID`. |

Use the ids already in the failing URL. Do not wrap them in another path.

```text
GRAPH_SITE_ID = {hostname}.sharepoint.com,{site-collection-id},{site-id}
GRAPH_LIST_ID = {list-id}
```

Copy both values from the failing list URL. Do not commit a real site id if this repository is public.

A successful body looks like `examples/aggregate-response.json`. The values in that file are examples, not a copy of a real list. Read these properties:

| Property | Meaning |
| --- | --- |
| `status` | `ok` means the function finished. It does not mean every item succeeded. |
| `elapsedMs` | Time spent inside the function. |
| `result.itemCount` | Rows in `result.items`. |
| `result.idCount` | Ids found before the item reads. |
| `result.batchCount` | How many Graph `$batch` calls were made. 1,500 items is 75. |
| `result.batchSize` | Always 20. |
| `result.failures` | Must be `[]`. Any entry is an item Graph did not return. |
| `result.items[].id` | SharePoint list item id. |
| `result.items[].OwnerLookupId` | Person column id, if you selected that column. A string, not a number. |
| `result.items[].Owner` | Display name of that same column. |

`mode=mock` is only a smoke test of the URL and the key. It does not prove SharePoint access.

## Create the Function App

Start on **Linux Basic B1**. Do not start on Consumption or Premium.

B1 is an always-on App Service plan. You pay for the plan hour, not for each call. One instance in East Asia is about $0.02 per hour on the public price list, about $15 per month, whether the function is idle or not. Delete or stop the plan to stop that charge. Storage and Application Insights are extra and are usually a few dollars at this volume. Your contract price can be lower. Check the price for your region before you create the plan.

A test of this pattern, 1,500 items, returned in about 54 seconds. That fits on B1. The HTTP response still has to come back within 230 seconds. A more expensive plan does not remove that limit.

Move to **Premium EP1** only if a B1 call is cut off, the first call after idle is too slow, or you need the function execution itself to run for more than a few minutes. EP1 keeps at least one instance allocated. In East Asia the public list price is about $0.20 per vCPU hour plus $0.0143 per GiB hour. EP1 is 1 vCPU and 3.5 GB, so one always-on instance is about $183 per month. Consumption is the wrong start: it can scale to zero, and one execution is capped at 10 minutes.

Sign in first. If `az account show` is the wrong subscription, run `az login` and `az account set --subscription YOUR-SUBSCRIPTION`. Creating the app does not deploy the code in this repo. The `aggregate` route appears only after the deploy step.

Replace `eastasia` if this workload lives in another region. `YOUR-FUNCTION-APP` and the storage account name must be globally unique.

### Portal

1. In the Azure portal, create a **Function App**.
2. Publish **Code**. Runtime stack **Python 3.11**. Operating system **Linux**. Version **4**.
3. Plan type **App Service**, operating system **Linux**, size **Basic B1**. Do not pick Consumption or Elastic Premium yet. Create the storage account the wizard asks for. Turn Application Insights on.
4. After deployment, open the app, then **Identity**, then **System assigned**, and set it to **On**. Save. Copy both ids:
   - **Object (principal) ID**
   - The application id. If the blade does not show it, run the command in the next section.
5. Open **Settings**, then **Environment variables**. Add:

   | Name | Value |
   | --- | --- |
   | `GRAPH_SITE_ID` | the site id above |
   | `GRAPH_LIST_ID` | `{list-id}` |

   Apply the change. The app restarts.

### CLI equivalent

```bash
az group create --name rg-sp-lookup-aggregate --location eastasia

az storage account create \
  --name stsplookupagg \
  --resource-group rg-sp-lookup-aggregate \
  --location eastasia \
  --sku Standard_LRS

az functionapp plan create \
  --name plan-sp-lookup-aggregate \
  --resource-group rg-sp-lookup-aggregate \
  --location eastasia \
  --sku B1 \
  --is-linux

az functionapp create \
  --name YOUR-FUNCTION-APP \
  --resource-group rg-sp-lookup-aggregate \
  --plan plan-sp-lookup-aggregate \
  --runtime python \
  --runtime-version 3.11 \
  --functions-version 4 \
  --storage-account stsplookupagg \
  --os-type Linux

az functionapp identity assign \
  --name YOUR-FUNCTION-APP \
  --resource-group rg-sp-lookup-aggregate

az functionapp config appsettings set \
  --name YOUR-FUNCTION-APP \
  --resource-group rg-sp-lookup-aggregate \
  --settings \
    GRAPH_SITE_ID="{hostname}.sharepoint.com,{site-collection-id},{site-id}" \
    GRAPH_LIST_ID="{list-id}"
```

`YOUR-FUNCTION-APP` must be globally unique. Change the storage account name if `stsplookupagg` is taken.

### Move to EP1 only if B1 is not enough

Do not change the B1 SKU to EP1. They are different plan types. Create a Premium plan and move the existing app. The identity, app settings, and deployed code stay with the app.

```bash
az functionapp plan create \
  --name plan-sp-lookup-ep1 \
  --resource-group rg-sp-lookup-aggregate \
  --location eastasia \
  --sku EP1 \
  --is-linux

az functionapp update \
  --name YOUR-FUNCTION-APP \
  --resource-group rg-sp-lookup-aggregate \
  --plan plan-sp-lookup-ep1
```

After the move, you can delete `plan-sp-lookup-aggregate` if no other app uses it. EP1 still does not remove the 230-second HTTP limit. If the caller times out there, lower `top` and call again.

## Use a coding agent

A coding agent can run the commands above for you. It does not get extra Azure rights. It can only do what `az login` can already do. The `Sites.Selected` grant still needs an account that can assign Graph application roles and SharePoint site permission.

Use the same plugin that was used to build this playbook: [azure-skills](https://github.com/microsoft/azure-skills), plugin `azure`. It installs the Azure skills and the [Azure MCP Server](https://github.com/Azure/mcp) (`npx -y @azure/mcp@latest server start`). The MCP server is what talks to the signed-in subscription. The skills tell the agent how a Function App, plan, and identity grant should be done.

In GitHub Copilot CLI, after `az login`:

```text
/plugin marketplace add microsoft/azure-skills
/plugin install azure@azure-skills
```

Then start `copilot` in this repository. The plugin registers the Azure MCP server. You do not need to add a second MCP server by hand. Confirm with `/mcp show`. The server should be `azure`, command `npx`, args `@azure/mcp@latest server start`. It uses the Azure CLI login. Do not put a client secret in the MCP config.

Open this repository and use a prompt like this. Do not paste a function key into a shared chat, and do not let the agent commit `local.settings.json` or a key.

```text
Deploy this repository as a Python 3.11 Linux Function App on Basic B1.
Turn on the system-assigned identity.
Set GRAPH_SITE_ID and GRAPH_LIST_ID from the list URL I paste.
Do not create Premium EP1 unless a B1 call exceeds 230 seconds or the cold start is a problem.
Do not print or commit the function key.
```

Then grant the identity, publish if the agent has not already, and call `mode=mock` with `top=2` before `mode=graph`.

## Let the function read the list

The function does not use the caller's Graph token. It signs in with its own managed identity. Two grants are required. Either one alone returns 403.

You need an identity that can assign app roles and grant site permission. That is a Privileged Role Administrator for the first command, and a SharePoint admin or an app with `Sites.FullControl.All` for the second.

```bash
APP_NAME="YOUR-FUNCTION-APP"
RG="rg-sp-lookup-aggregate"
SITE_ID="{hostname}.sharepoint.com,{site-collection-id},{site-id}"

PRINCIPAL_ID=$(az functionapp identity show --name "$APP_NAME" --resource-group "$RG" --query principalId -o tsv)
APP_ID=$(az ad sp show --id "$PRINCIPAL_ID" --query appId -o tsv)
GRAPH_SP_ID=$(az ad sp list --filter "appId eq '00000003-0000-0000-c000-000000000000'" --query "[0].id" -o tsv)
ROLE_ID=$(az ad sp list --filter "appId eq '00000003-0000-0000-c000-000000000000'" --query "[0].appRoles[?value=='Sites.Selected'].id | [0]" -o tsv)

az rest --method POST \
  --uri "https://graph.microsoft.com/v1.0/servicePrincipals/${PRINCIPAL_ID}/appRoleAssignments" \
  --body "{\"principalId\":\"${PRINCIPAL_ID}\",\"resourceId\":\"${GRAPH_SP_ID}\",\"appRoleId\":\"${ROLE_ID}\"}"

az rest --method POST \
  --uri "https://graph.microsoft.com/v1.0/sites/${SITE_ID}/permissions" \
  --body "{\"roles\":[\"read\"],\"grantedToIdentities\":[{\"application\":{\"id\":\"${APP_ID}\",\"displayName\":\"${APP_NAME}\"}}]}"
```

`Sites.Selected` is the Microsoft Graph application role `883ea226-0bf2-4a8f-9f9d-92c9162a727d`. The second call must use `APP_ID`, the application id, not the object id. Role `read` is enough. It does not change the list.

If this app already has `Sites.Read.All` and an admin has consented it, you can skip the site permission call. `Sites.Selected` plus one site grant is the narrower choice.

## Deploy this repo

Install [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local) v4, then from this folder:

```bash
func azure functionapp publish YOUR-FUNCTION-APP --python
```

That builds `requirements.txt` on Azure. The first publish takes a few minutes. When it finishes, the portal shows a function named `aggregate`.

If publish stops while waiting for deployment settings, pack the folder and use remote build:

```bash
zip -r /tmp/aggregate.zip function_app.py host.json requirements.txt -x "*.pyc"
az functionapp deployment source config-zip \
  --name YOUR-FUNCTION-APP \
  --resource-group rg-sp-lookup-aggregate \
  --src /tmp/aggregate.zip \
  --build-remote true
```

Do not commit `local.settings.json`. It is for an optional local run only. Copy `local.settings.json.example` if you want `func start` on your machine. Production uses the app settings from the previous section, not that file.

## Call it

First prove the route with two synthetic rows. This does not touch SharePoint.

```bash
curl --request POST \
  "https://YOUR-FUNCTION-APP.azurewebsites.net/api/aggregate" \
  --header "Content-Type: application/json" \
  --header "x-functions-key: YOUR-FUNCTION-KEY" \
  --data '{"mode":"mock","top":2,"fields":["Title","OwnerLookupId","Owner"]}'
```

Expect HTTP 200, `result.itemCount` of 2, and non-null `Owner` values.

Get the key in the portal: open the Function App, open **Functions**, open **aggregate**, then **Function keys**. Copy the `default` key. Treat it as a password. If it is pasted into email or chat, generate a new key and delete the old one.

Then send the real request. This is the body to use:

```bash
curl --request POST \
  "https://YOUR-FUNCTION-APP.azurewebsites.net/api/aggregate" \
  --header "Content-Type: application/json" \
  --header "x-functions-key: YOUR-FUNCTION-KEY" \
  --max-time 240 \
  --data @- <<'JSON'
{
  "mode": "graph",
  "top": 1500,
  "fields": [
    "Title",
    "Owner",
    "OwnerLookupId",
    "Category",
    "CategoryLookupId"
  ]
}
JSON
```

Start with `"top": 5` if you only want to confirm permissions. A 1,500-item call is supposed to take about a minute. The first call after the app has been idle can take longer.

### Optional: send the same call from Bruno

curl is enough. If you already use Bruno, this repo also contains the same request.

1. Install [Bruno](https://www.usebruno.com/).
2. Open the collection in `bruno/Graph Aggregate`.
3. Copy `environments/Example.yml.example` to `environments/Example.yml`.
4. Set `functionBase` to `https://YOUR-FUNCTION-APP.azurewebsites.net` with no trailing slash and no `/api`.
5. Paste the function key into `functionKey`.
6. Select the **Example** environment.
7. Open **Aggregate** and send it.

The request URL is `{{functionBase}}/api/aggregate?code={{functionKey}}`. The body is the same JSON as the curl example above. The tests expect HTTP 200, 1,500 items, `batchSize` 20, an empty `failures` array, and a non-null `OwnerLookupId` plus display name on the first row. Do not commit `Example.yml`.

## Limits

| Limit | Value | What happens if you cross it |
| --- | --- | --- |
| Graph `$batch` size | 20 | HTTP 400. The code already splits at 20. |
| Batches in flight | 3 | More than this is more likely to receive HTTP 429. |
| `top` | 2000 | The function rejects a larger value by clamping. Ask for 1500. |
| HTTP response | 230 seconds | The caller sees a timeout even if the plan allows a longer function. Lower `top`. |
| Graph 429 | `Retry-After` | The function retries the batch up to 5 times. A still-failing item is listed in `result.failures`, not silently dropped. |

Column internal names are case-sensitive. A misspelled name comes back as `null` for that property only. That is not the 12-column bug.

## Troubleshooting

| Result | Cause | Fix |
| --- | --- | --- |
| 401 | Missing or wrong function key | Send `x-functions-key`, or `?code=`. |
| 400 `siteId and listId` | App settings missing | Save `GRAPH_SITE_ID` and `GRAPH_LIST_ID`, wait for the restart, call again. |
| 403 from Graph | Identity cannot read this site | Confirm `Sites.Selected` is assigned, then confirm the site grant used the application id. |
| 502 | Graph rejected the read | Read `detail` in the response. Do not retry in a loop. |
| HTTP 200 and a non-empty `failures` | One or more item reads failed | Those ids are not in `items`. If `status` is 429, wait and call again. |
| Timeout around 230 seconds | Too much throttling for one synchronous call | Call with a smaller `top`. Do not raise the batch size. |
| A selected column is null on a successful row | The column is empty on that item, or the internal name is wrong | Compare with one direct `GET /items/{id}`. |

Logs are in the Function App, under **Log stream**, or in Application Insights, table `requests` and `traces`.
