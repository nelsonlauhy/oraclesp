let accessToken = null;
let currentDriveId = null;
let currentFolderId = null;
let breadcrumb = [];

const msalConfig = {
  auth: {
    clientId: "31109757-b52b-4b7c-96c5-454490ad4c4e", // ✅ Your app’s client ID
    authority: "https://login.microsoftonline.com/12551c0b-1ff7-4427-accf-80d5406276e0", // ✅ Your tenant ID
    redirectUri: window.location.href
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function signIn() {
  try {
    const result = await msalInstance.loginPopup({
      scopes: ["Sites.Read.All", "Files.Read.All"]
    });

    accessToken = result.accessToken;
    document.getElementById("librarySelect").disabled = false;
    loadLibraries();
  } catch (err) {
    console.error("Login failed", err);
  }
}

async function loadLibraries() {
  const siteId = "oraclegrouprealty.sharepoint.com,8b247071-c01e-4b83-958e-413d2e156b40,7892dd7a-2e84-4201-a4d4-dba0582d500e";
  const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  const data = await res.json();
  const select = document.getElementById("librarySelect");
  select.innerHTML = "";
  data.value.forEach(lib => {
    const option = document.createElement("option");
    option.value = lib.id;
    option.textContent = lib.name;
    select.appendChild(option);
  });

  select.addEventListener("change", () => {
    currentDriveId = select.value;
    breadcrumb = [];
    loadFiles(currentDriveId);
  });
}

async function loadFiles(driveId, folderId = "root") {
  currentFolderId = folderId;
  showLoading();

  const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  const data = await res.json();
  const list = document.getElementById("fileList");
  list.innerHTML = "";

  updateBreadcrumb();

  data.value.forEach(item => {
    const a = document.createElement("a");
    a.className = "list-group-item list-group-item-action d-flex justify-content-between align-items-center";
    a.innerHTML = `<span><i class="bi ${item.folder ? 'bi-folder' : 'bi-file-earmark'} me-2"></i>${item.name}</span>`;
    if (item.folder) {
      a.href = "#";
      a.onclick = () => {
        breadcrumb.push({ id: item.id, name: item.name });
        loadFiles(driveId, item.id);
        return false;
      };
    } else {
      a.href = item.webUrl;
      a.target = "_blank";
    }
    list.appendChild(a);
  });
}

function updateBreadcrumb() {
  const breadcrumbEl = document.getElementById("breadcrumb");
  breadcrumbEl.innerHTML = "";
  const rootCrumb = document.createElement("li");
  rootCrumb.className = "breadcrumb-item";
  rootCrumb.innerHTML = `<a href="#">Root</a>`;
  rootCrumb.onclick = () => {
    breadcrumb = [];
    loadFiles(currentDriveId);
  };
  breadcrumbEl.appendChild(rootCrumb);

  breadcrumb.forEach((crumb, index) => {
    const li = document.createElement("li");
    li.className = "breadcrumb-item";
    li.innerHTML = `<a href="#">${crumb.name}</a>`;
    li.onclick = () => {
      breadcrumb = breadcrumb.slice(0, index + 1);
      loadFiles(currentDriveId, crumb.id);
    };
    breadcrumbEl.appendChild(li);
  });
}

function showLoading() {
  document.getElementById("fileList").innerHTML = `
    <div class="text-center py-3">
      <div class="spinner-border text-primary" role="status"></div>
    </div>
  `;
}
