const msalConfig = {
  auth: {
    clientId: "31109757-b52b-4b7c-96c5-454490ad4c4e",
    authority: "https://login.microsoftonline.com/12551c0b-1ff7-4427-accf-80d5406276e0",
    redirectUri: window.location.href
  }
};

const siteId = "oraclegrouprealty.sharepoint.com,8b247071-c01e-4b83-958e-413d2e156b40,7892dd7a-2e84-4201-a4d4-dba0582d500e";
const msalInstance = new msal.PublicClientApplication(msalConfig);

let accessToken = null;
let currentDriveId = null;
let currentFolderId = null;
let breadcrumb = [];
let selectedFileItems = [];

window.onload = async () => {
  const currentAccounts = msalInstance.getAllAccounts();
  if (currentAccounts.length > 0) {
    try {
      const result = await msalInstance.acquireTokenSilent({
        scopes: ["Sites.Read.All", "Files.Read.All", "Mail.ReadWrite"],
        account: currentAccounts[0]
      });
      accessToken = result.accessToken;
      updateUIAfterLogin(currentAccounts[0]);
      document.getElementById("librarySelect").disabled = false;
      loadLibraries();
    } catch (error) {
      console.warn("Silent token acquisition failed", error);
    }
  }
};

async function signIn() {
  try {
    const result = await msalInstance.loginPopup({
      scopes: ["Sites.Read.All", "Files.Read.All", "Mail.ReadWrite"]
    });

    accessToken = result.accessToken;
    updateUIAfterLogin(result.account);
    document.getElementById("librarySelect").disabled = false;
    loadLibraries();
  } catch (err) {
    console.error("Login failed", err);
  }
}

function updateUIAfterLogin(account) {
  document.getElementById("signInBtn").style.display = "none";
  document.getElementById("welcomeMessage").style.display = "none";
  document.getElementById("userStatus").style.display = "inline-block";
  document.getElementById("userStatus").textContent = `✅ Signed in as ${account.username}`;
}

async function loadLibraries() {
  const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  const data = await res.json();
  const select = document.getElementById("librarySelect");
  select.innerHTML = "";

  const sortedLibraries = data.value.sort((a, b) => a.name.localeCompare(b.name));

  sortedLibraries.forEach(lib => {
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

  if (select.options.length > 1) {
    select.selectedIndex = 1;
    select.dispatchEvent(new Event('change'));
  }
}

function updateSelectedTags() {
  const tagBox = document.getElementById("selectedTags");
  tagBox.innerHTML = "";
  selectedFileItems.forEach(file => {
    const span = document.createElement("span");
    span.className = "badge bg-primary d-flex align-items-center";
    span.innerHTML = `${file.name} <button type="button" class="btn-close btn-close-white btn-sm ms-2" aria-label="Remove" onclick="removeFileFromSelection('${file.itemId}')"></button>`;
    tagBox.appendChild(span);
  });
}

function addFileToSelection(itemId, driveId, name, size) {
  const exists = selectedFileItems.some(f => f.itemId === itemId);
  if (!exists) {
    selectedFileItems.push({ driveId, itemId, name, size });
    updateSelectedTags();
  }
}

function removeFileFromSelection(itemId) {
  selectedFileItems = selectedFileItems.filter(f => f.itemId !== itemId);
  updateSelectedTags();
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
    const icon = item.folder ? 'bi-folder' : 'bi-file-earmark';
    const a = document.createElement("a");
    a.className = "list-group-item list-group-item-action";

    if (item.folder) {
      a.innerHTML = `<div class="file-row" style="border-bottom: none;"><div class="file-name"><i class="bi ${icon} me-2"></i><strong>${item.name}</strong></div></div>`;
      a.href = "#";
      a.onclick = () => {
        breadcrumb.push({ id: item.id, name: item.name });
        loadFiles(driveId, item.id);
        return false;
      };
    } else {
      a.innerHTML = `
        <div class="file-row">
          <div class="file-name"><button class="btn btn-sm btn-outline-success me-2" onclick="event.preventDefault(); addFileToSelection('${item.id}', '${driveId}', '${item.name}', ${item.size || 0})"><i class="bi bi-plus-lg"></i></button><i class="bi ${icon} me-2"></i>${item.name}
          </div>
          <div class="file-action">
            <button class="btn btn-sm btn-outline-success" onclick="event.preventDefault(); addFileToSelection('${item.id}', '${driveId}', '${item.name}', ${item.size || 0})">
              <i class="bi bi-plus-lg"></i>
            </button>
          </div>
        </div>
      `;
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

async function submitFiles() {
  const modalBody = document.getElementById("modalBody");
  const createBtn = document.getElementById("createEmailBtn");

  modalBody.innerHTML = "";

  if (selectedFileItems.length === 0) {
    modalBody.innerHTML = `<div class="text-muted">No files selected.</div>`;
    createBtn.disabled = true;
  } else {
    const list = document.createElement("ul");
    list.className = "list-group";

    selectedFileItems.forEach(file => {
      const li = document.createElement("li");
      li.className = "list-group-item";
      li.textContent = file.name;
      list.appendChild(li);
    });

    modalBody.appendChild(list);
    createBtn.disabled = false;
  }

  const fileModal = new bootstrap.Modal(document.getElementById("fileModal"));
  fileModal.show();
}

async function createDraftEmailWithAttachments() {
  const modalBody = document.getElementById("modalBody");
  const createBtn = document.getElementById("createEmailBtn");
  const cancelBtn = document.querySelector("#fileModal .btn-secondary");

  createBtn.disabled = true;
  modalBody.innerHTML += `<div class="text-info mt-3">📥 Downloading files and creating draft email...</div>`;

  try {
    const emailRes = await fetch("https://graph.microsoft.com/v1.0/me/messages", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        subject: "New auto email attachment",
        body: {
          contentType: "HTML",
          content: "<p>This email was generated automatically with attachments.</p>"
        },
        toRecipients: []
      })
    });

    const emailData = await emailRes.json();
    const messageId = emailData.id;

    for (const file of selectedFileItems) {
      const base64data = await downloadFileAsBase64(file.driveId, file.itemId);

      await fetch(`https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: file.name,
          contentBytes: base64data,
          contentType: "application/octet-stream"
        })
      });
    }

    modalBody.innerHTML += `<div class="text-success mt-3">✅ Draft email with attachments created.<br><a href="https://outlook.office365.com/mail/drafts" target="_blank">Open Drafts</a></div>`;
    cancelBtn.textContent = "Close";
    createBtn.disabled = true;
  } catch (error) {
    console.error("Error creating draft email:", error);
    modalBody.innerHTML += `<div class="text-danger mt-3">❌ Failed to create draft email.</div>`;
    createBtn.disabled = false;
  }
}

async function downloadFileAsBase64(driveId, itemId) {
  const response = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  const blob = await response.blob();

  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const base64data = reader.result.split(',')[1];
      resolve(base64data);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
