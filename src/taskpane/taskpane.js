/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("login-btn").addEventListener("click", login);
  document.getElementById("logout-btn").addEventListener("click", logout);
  checkLoginStatus();
});

const apiBase = "https://muslimxchange.com/wp-json/mx/v1";

// 🔁 Replace IndexedDB with OfficeRuntime.storage
async function saveToken(token) {
  await OfficeRuntime.storage.setItem("jwt", token);
}

async function getToken() {
  return await OfficeRuntime.storage.getItem("jwt");
}

async function clearToken() {
  await OfficeRuntime.storage.removeItem("jwt");
}

async function checkLoginStatus() {
  const token = await getToken();
  const loginSection = document.getElementById("login-section");
  const statusDiv = document.getElementById("login-status");
  const logoutBtn = document.getElementById("logout-btn");

  if (token) {
    try {
      const res = await fetch(`${apiBase}/secure-hello`, {
        headers: { Authorization: `Bearer ${token}` }
      });

      if (res.ok) {
        const data = await res.json();
        loginSection.style.display = "none";
        statusDiv.innerText = `✅ Logged in as ${data.user.name}`;
        logoutBtn.style.display = "inline-block";
      }
    } catch (err) {
      console.error("Error checking login status", err);
    }
  }
}

async function login() {
  const username = document.getElementById("username").value.trim();
  const password = document.getElementById("password").value;
  const statusDiv = document.getElementById("login-status");
  const loginSection = document.getElementById("login-section");
  const logoutBtn = document.getElementById("logout-btn");

  statusDiv.innerText = "🔄 Logging in...";

  try {
    const res = await fetch(`${apiBase}/login`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, password }),
    });

    const json = await res.json();

    if (!res.ok) {
      console.error("Login failed:", json);
      statusDiv.innerText = `❌ ${json.message || "Login failed."}`;
      return;
    }

    const { token, user } = json;

    await saveToken(token);
    loginSection.style.display = "none";
    statusDiv.innerText = `✅ Logged in as ${user.name}`;
    logoutBtn.style.display = "inline-block";

  } catch (err) {
    console.error("Login error:", err);
    statusDiv.innerText = "❌ Network or CORS error during login.";
  }
}

async function logout() {
  await clearToken();
  document.getElementById("login-section").style.display = "block";
  document.getElementById("logout-btn").style.display = "none";
  document.getElementById("login-status").innerText = "Logged out.";
}

async function callSecureAPI() {
  const responseDiv = document.getElementById("api-response");
  const token = await getToken();

  if (!token) {
    responseDiv.innerText = "❌ No token found. Please log in.";
    return;
  }

  try {
    const res = await fetch(`${apiBase}/secure-hello`, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!res.ok) {
      responseDiv.innerText = "❌ Secure call failed.";
      return;
    }

    const data = await res.json();
    responseDiv.innerText = `✅ ${data.message} from ${data.user.name}`;
  } catch (err) {
    responseDiv.innerText = "❌ API request error.";
    console.error(err);
  }
}

// (Optional — if you're keeping it in taskpane.js too)
async function checkCompliant(ticker) {
  try {
    const res = await fetch(`${apiBase}/hello?ticker=${encodeURIComponent(ticker)}`);
    if (!res.ok) return "❌ Error";

    const data = await res.json();
    return data.message || `✅ ${ticker} Compliant`;
  } catch (err) {
    console.error("Compliance check failed:", err);
    return "❌ API Error";
  }
}

