MuslimXchange Excel Add-in
The MuslimXchange Excel Add-in provides custom functions to fetch Islamic compliance status and financial data directly from MuslimXchange.com. It supports three functions: COMPLIANT(ticker), TICKER(ticker, fields[]), and ISIN(isin, fields[]), all powered by a secure WordPress REST API.

A valid JWT token is required, stored using OfficeRuntime.storage. The API base URL is https://muslimxchange.com/wp-json/mx/v1. Without authentication, functions will return a "Login required" message.

