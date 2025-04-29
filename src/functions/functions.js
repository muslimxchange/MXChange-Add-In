/* eslint-disable @typescript-eslint/no-unused-vars */
/* global console setInterval, clearInterval */
/**
 * Excel custom function: Checks if a ticker is compliant via WordPress API.
 * @customfunction
 * @param {string} ticker The ticker symbol to check.
 * @returns {Promise<string>} The compliance status message.
 */
/**
 * @customfunction COMPLIANT
 * @param {string} ticker
 * @returns {Promise<string>}
 */
async function Compliant(ticker) {
  const apiBase = "https://muslimxchange.com/wp-json/mx/v1";

  try {
    const token = await OfficeRuntime.storage.getItem("jwt");

    if (!token) {
      return [["Login required"]];
    }
    const callingURL = `${apiBase}/hello?ticker=${encodeURIComponent(ticker)}&token=${encodeURIComponent(token || "")}`;

    let res;
    try {
      res = await fetch(callingURL);
    } catch {
      return "Network Error";
    }

    if (!res.ok) {
      return `API Error ${res.status}`;
    }

    const data = await res.json();
    return data.message || `Unknown status for ${ticker}`;
  } catch {
    return "Request Failed";
  }
}

/**
 * Fetches data for a ticker with specific fields.
 * @customfunction TICKER
 * @param {string} ticker
 * @param {string[]} fields
 * @returns {Promise<any[][]>}
 */
async function TICKER(ticker, ...fields) {
  const apiBase = "https://muslimxchange.com/wp-json/mx/v1";
  const token = await OfficeRuntime.storage.getItem("jwt");
  
  if (!token) {
    return [["Login required"]];
  }
  const fieldList = fields.flat().join(",");

  const url = `${apiBase}/ticker-data?ticker=${encodeURIComponent(ticker)}&fields=${encodeURIComponent(fieldList)}&token=${encodeURIComponent(token || "")}`;

  try {
    const res = await fetch(url);
    if (!res.ok) return [[`API Error ${res.status}`]];

    const data = await res.json();
    const row = [];

    const flatFields = fields.flat();
    for (let i = 0; i < flatFields.length; i++) {
      const field = flatFields[i];
      const value = data[field];
      const isLast = i === flatFields.length - 1;
    
      // Push only if it's not the last OR it's defined
      if (!isLast || (value !== undefined && value !== null && value !== "")) {
        row.push(value ?? "Unknown");
      }
    }
    
    return [row];
  } catch {
    return [["Request Failed"]];
  }
}


/**
 * Fetches data for a ISIN with specific fields.
 * @customfunction ISIN
 * @param {string} isin
 * @param {string[]} fields
 * @returns {Promise<any[][]>}
 */
async function ISIN(ISIN, ...fields) {
  const apiBase = "https://muslimxchange.com/wp-json/mx/v1";
  const token = await OfficeRuntime.storage.getItem("jwt");
  
  if (!token) {
    return [["Login required"]];
  }
  const fieldList = fields.flat().join(",");

  const url = `${apiBase}/isin-data?isin=${encodeURIComponent(ISIN)}&fields=${encodeURIComponent(fieldList)}&token=${encodeURIComponent(token || "")}`;

  try {
    const res = await fetch(url);
    if (!res.ok) return [[`API Error ${res.status}`]];

    const data = await res.json();
    const row = [];

    const flatFields = fields.flat();
    for (let i = 0; i < flatFields.length; i++) {
      const field = flatFields[i];
      const value = data[field];
      const isLast = i === flatFields.length - 1;
    
      // Push only if it's not the last OR it's defined
      if (!isLast || (value !== undefined && value !== null && value !== "")) {
        row.push(value ?? "Unknown");
      }
    }
    
    return [row];
  } catch {
    return [["Request Failed"]];
  }
}



/**
 * Appends a message to the log.
 * @customfunction LOG
 * @param {string} message String to log.
 * @returns {Promise<string>} The logged message.
 
async function logMessage(message) {
  let existingLog = await OfficeRuntime.storage.getItem("cfLogs").catch(() => "");
  let updatedLog = (existingLog || "") + message + "\n";
  await OfficeRuntime.storage.setItem("cfLogs", updatedLog);
  return message;
}

/**
 * Returns all logged messages as a vertical array.
 * @customfunction GETLOGS
 * @returns {Promise<string[][]>} Array of log entries, one per row.
 
async function getLogs() {
  let log = await OfficeRuntime.storage.getItem("cfLogs").catch(() => "");
  let lines = (log || "").trim().split("\n");
  return lines.map(line => [line]); // Make it 2D vertical array
}
*/
