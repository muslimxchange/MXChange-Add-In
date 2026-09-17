# MXChange add-in 1.3.3

Requires MX API 1.9.0. Publish these files to the root of the `MXChange-Add-In` GitHub Pages repo and reload the 1.3.1.0 manifest in Excel. The manifest GUID and the `MX.TICKER`, `MX.ISIN` and `MX.FIELDS` signatures are unchanged; the `MX.TICKERIN` / `MX.ISININ` functions from the unreleased 1.3.0 are not published.

- Tickers are exchange-suffixed in the dataset (ASML, ASML.AS), so `MX.TICKER` always resolves to one listing.
- An ISIN names the security, not the listing. `MX.ISIN` returns **one row per listing** (a dual-listed ISIN spills two rows); add `"ticker"` or `"Market"` to the fields to see which row is which, e.g. `=MX.ISIN("NL0010273215","ticker","Market","Result")`.
- Three shared workers per runtime cover all batches and `MX.FIELDS`; a 20-second deadline includes the response body and also applies to sign-in.
- Sign-out invalidates retries and stale results; credentials are stored in `mx_session_v2`, and a sign-out marker overrides any legacy `mx_auth`.
- Fields outside your plan show `Not in your plan: <field>` in that cell; `MX.FIELDS()` now lists tier and availability per field. Trial accounts see their daily allowance in the task pane.
- The server has no batch cap; the add-in sends up to 1,000 identifiers per request so responses stay quick.
