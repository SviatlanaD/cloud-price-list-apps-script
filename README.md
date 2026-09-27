# Cloud price list builder (Google Apps Script)

Builds a multi-currency price list for Azure virtual machines inside Google
Sheets and exports it as an `.xlsx` file into a chosen Drive folder.

Both data sources are public and need no API key:

* [Azure Retail Prices API](https://learn.microsoft.com/en-us/rest/api/cost-management/retail-prices/azure-retail-prices) — VM prices per region, in USD
* [Frankfurter](https://frankfurter.dev/) — European Central Bank reference rates

## What it does

```
Azure API  ──┐
             ├──►  price_list sheet  ──►  .xlsx in a Drive folder
ECB rates  ──┘
```

Four steps, available from a **Price list** menu that appears in the
spreadsheet:

1. **Import prices** — pulls prices for one service in one region, follows the
   API paging, drops Spot and Low Priority meters
2. **Import FX rates** — pulls ECB rates and rebases them from EUR onto USD,
   since Azure quotes in USD
3. **Build price list** — joins the two and adds hourly and monthly columns per
   currency (monthly = 730 hours)
4. **Export .xlsx to Drive** — saves the result into the folder from the config

## Setup

1. Create a new Google Sheet.
2. Add a sheet named `config` with keys in column A and values in column B:

   | | A | B |
   |---|---|---|
   |1| `service` | `Virtual Machines` |
   |2| `region` | `westeurope` |
   |3| `currencies` | `USD, EUR, PLN, GBP, CZK` |
   |4| `max_rows` | `200` |
   |5| `drive_folder_id` | id of the target folder |

   The folder id is the last part of its URL:
   `https://drive.google.com/drive/folders/`**`1AbC...`**

3. **Extensions → Apps Script**, paste `Code.gs`, save.
4. Reload the spreadsheet. A **Price list** menu appears.
5. First run asks for permission — the script needs to call two external APIs
   and write a file to Drive.

The other sheets (`prices_raw`, `fx_rates`, `price_list`) are created
automatically.

## The .xlsx export

This is the part worth explaining. Apps Script has no direct "save as xlsx",
and the spreadsheet export endpoint ignores the `gid` parameter for xlsx — it
always returns the whole workbook. So the script copies the price list into a
throwaway spreadsheet, exports that one, writes the blob into the target folder
and trashes the temporary file:

```js
var url = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
var blob = UrlFetchApp.fetch(url, {
  headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
}).getBlob().setName(fileName);
folder.createFile(blob);
```

Files are named `price_list_<region>_<yyyy-MM-dd_HHmm>.xlsx`.

## Notes

* Spot and Low Priority meters are filtered out. They are priced on a different
  basis and would not be comparable with pay-as-you-go rates.
* 730 hours is the convention cloud vendors use for a month, not an exact month.
* ECB publishes rates once per working day, so a price list built on a weekend
  carries Friday's rates. The date is kept in the `fx_rates` sheet.
* `max_rows` caps the import. Without it a broad service and region can return
  thousands of meters.
