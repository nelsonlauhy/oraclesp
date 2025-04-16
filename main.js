const msalConfig = {
    auth: {
      clientId: "31109757-b52b-4b7c-96c5-454490ad4c4e", // ✅ Your app’s client ID
      authority: "https://login.microsoftonline.com/12551c0b-1ff7-4427-accf-80d5406276e0", // ✅ Your tenant ID
      redirectUri: window.location.href
    }
  };
  
  const msalInstance = new msal.PublicClientApplication(msalConfig);
  let accessToken = null;
  
  async function signIn() {
    try {
      const result = await msalInstance.loginPopup({
        scopes: ["Sites.Read.All", "Files.Read.All"]
      });
  
      accessToken = result.accessToken || (await msalInstance.acquireTokenSilent({
        scopes: ["Sites.Read.All", "Files.Read.All"],
        account: result.account
      })).accessToken;
  
      document.getElementById("librarySelect").disabled = false;
      loadLibraries();
    } catch (err) {
      console.error(err);
    }
  }
  
  async function loadLibraries() {
    const siteId = "oraclegrouprealty.sharepoint.com,8b247071-c01e-4b83-958e-413d2e156b40,7892dd7a-2e84-4201-a4d4-dba0582d500e";

  
    const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    });
  
    const data = await response.json();
    const select = document.getElementById("librarySelect");
    select.innerHTML = "";
    data.value.forEach(lib => {
      const option = document.createElement("option");
      option.value = lib.id;
      option.textContent = lib.name;
      select.appendChild(option);
    });
  
    select.addEventListener("change", () => loadFiles(select.value));
  }
  
  async function loadFiles(driveId) {
    const response = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`, {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    });
  
    const data = await response.json();
    const list = document.getElementById("fileList");
    list.innerHTML = "";
    data.value.forEach(file => {
      const item = document.createElement("a");
      item.href = file.webUrl;
      item.className = "list-group-item list-group-item-action";
      item.target = "_blank";
      item.textContent = file.name;
      list.appendChild(item);
    });
  }
  