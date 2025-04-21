const msalConfig = {
  auth: {
    clientId: "31109757-b52b-4b7c-96c5-454490ad4c4e",
    authority: "https://login.microsoftonline.com/12551c0b-1ff7-4427-accf-80d5406276e0",
    redirectUri: window.location.href
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

let accessToken = null;
let currentSiteId = null;
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

      // Activate dropdown after login
      document.getElementById("siteSelect").onchange = () => {
        currentSiteId = document.getElementById("siteSelect").value;
        loadLibraries();
      };
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

    // Activate dropdown after login
    document.getElementById("siteSelect").onchange = () => {
      currentSiteId = document.getElementById("siteSelect").value;
      loadLibraries();
    };
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
  // Hide main section by default when a new site is selected
  document.getElementById("mainContentSection").classList.add("d-none");
  document.getElementById("fileList").innerHTML = "";
  document.getElementById("selectedTags").innerHTML = "";

  const select = document.getElementById("librarySelect");
  select.innerHTML = '<option selected disabled>Select a Document Library</option>';
  select.disabled = true;

  try {
    const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/drives`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!res.ok) {
      throw new Error(`Access denied or error loading drives (status: ${res.status})`);
    }

    const data = await res.json();

    const sortedLibraries = data.value.sort((a, b) => a.name.localeCompare(b.name));
    sortedLibraries.forEach(lib => {
      const option = document.createElement("option");
      option.value = lib.id;
      option.textContent = lib.name;
      select.appendChild(option);
    });

    select.disabled = false;

    select.onchange = () => {
      currentDriveId = select.value;
      breadcrumb = [];
      loadFiles(currentDriveId);

      const selectedLibraryText = select.options[select.selectedIndex].textContent;
      if (selectedLibraryText !== "Select a Document Library") {
        document.getElementById("mainContentSection").classList.remove("d-none");
      }
    };

    // Auto-select first library if available
    if (select.options.length > 1) {
      select.selectedIndex = 1;
      select.dispatchEvent(new Event("change"));
    }
  } catch (error) {
    console.error("Error loading libraries:", error);
    // Optional: show a user-friendly alert
    alert("⚠️ You do not have permission to access this SharePoint site.");
  }
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
    a.className = "list-group-item list-group-item-action border-0";

    if (item.folder) {
      a.innerHTML = `<div class="file-row"><div class="file-name"><i class="bi ${icon} me-2"></i><strong>${item.name}</strong></div></div>`;
      a.href = "#";
      a.onclick = () => {
        breadcrumb.push({ id: item.id, name: item.name });
        loadFiles(driveId, item.id);
        return false;
      };
    } else {
      let approvalButtonHTML = '';
      if (item.name.toLowerCase().endsWith('.pdf')) {
        approvalButtonHTML = `
          <button class="btn btn-sm btn-add-lightgrey me-2" title="Request Document Approval"
            onclick="event.preventDefault(); openApprovalModal('${item.id}', '${driveId}', '${item.name}')">
            <i class="bi bi-check-lg"></i>
          </button>`;

      }

      a.innerHTML = `
        <div class="file-row">
          <div class="file-name">
            <button class="btn btn-sm btn-add-lightgrey me-1" title="Send as Email Attachment"
              onclick="event.preventDefault(); addFileToSelection('${item.id}', '${driveId}', '${item.name}', ${item.size || 0})">
              <i class="bi bi-plus-lg"></i>
            </button>
            ${approvalButtonHTML}
            <i class="bi ${icon} me-2"></i>${item.name}
          </div>
        </div>
      `;

      a.href = item.webUrl;
      a.target = "_blank";
    }

    list.appendChild(a);
  });

  if (data.value.length >= 0) {
    document.getElementById("mainContentSection").classList.remove("d-none");
  }
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

async function submitFiles() {
  const modalFileList = document.getElementById("modalFileList");
  const modalTotalSize = document.getElementById("modalTotalSize");
  const createBtn = document.getElementById("createEmailBtn");

  modalFileList.innerHTML = "";
  modalTotalSize.innerHTML = "";

  if (selectedFileItems.length === 0) {
    modalFileList.innerHTML = `<li class="list-group-item text-muted">No files selected.</li>`;
    createBtn.disabled = true;
  } else {
    let totalSize = 0;

    selectedFileItems.forEach(file => {
      const li = document.createElement("li");
      li.className = "list-group-item d-flex justify-content-between align-items-center";
      const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
      li.innerHTML = `<span>${file.name}</span><span class="text-muted">${sizeMB} MB</span>`;
      modalFileList.appendChild(li);
      totalSize += file.size;
    });

    const totalSizeMB = (totalSize / (1024 * 1024)).toFixed(2);
    modalTotalSize.textContent = `📦 Total Size: ${totalSizeMB} MB`;

    if (totalSize > 25 * 1024 * 1024) {
      modalTotalSize.innerHTML += `<br><span class="text-danger">⚠️ Total size exceeds 25MB limit</span>`;
      createBtn.disabled = true;
    } else {
      createBtn.disabled = false;
    }
  }

  const fileModal = new bootstrap.Modal(document.getElementById("fileModal"));
  fileModal.show();
}

async function createDraftEmailWithAttachments() {
  const modalBody = document.getElementById("modalBody");
  const createBtn = document.getElementById("createEmailBtn");
  const cancelBtn = document.querySelector("#fileModal .btn-secondary");

  createBtn.disabled = true;
  modalBody.innerHTML += `<div class="text-info mt-3">📥 Creating draft email and uploading files...</div>`;

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

    modalBody.innerHTML += `<div class="text-success mt-3">✅ Draft email created.<br><a href="https://outlook.office365.com/mail/drafts" target="_blank">Open Drafts</a></div>`;
    cancelBtn.textContent = "Close";
  } catch (error) {
    console.error("Error creating draft email:", error);
    modalBody.innerHTML += `<div class="text-danger mt-3">❌ Failed to create draft email.</div>`;
    createBtn.disabled = false;
  }
}

async function downloadFileAsBase64(driveId, itemId) {
  const response = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`, {
    headers: { Authorization: `Bearer ${accessToken}` }
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

async function openApprovalModal(itemId, driveId, fileName) {
  const modal = new bootstrap.Modal(document.getElementById('approvalModal'));
  const approverSelect = document.getElementById('approverSelect');
  const confirmBtn = document.getElementById('confirmApprovalBtn');

  // Reset dropdown and button
  approverSelect.innerHTML = '<option disabled selected>Loading recent contacts...</option>';
  confirmBtn.disabled = true;

  try {
    // ✅ Use /me/people instead of /users
    const res = await fetch("https://graph.microsoft.com/v1.0/me/people?$select=displayName,emailAddresses&$top=20", {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    const data = await res.json();
    console.log("People response:", data);

    if (!data || !data.value || data.value.length === 0) {
      approverSelect.innerHTML = '<option disabled selected>No contacts found</option>';
      return;
    }

    // Build dropdown
    approverSelect.innerHTML = '<option disabled selected>Select an approver</option>';
    data.value.forEach(person => {
      const email = person.emailAddresses?.[0]?.address;
      if (email) {
        const option = document.createElement("option");
        option.value = email;
        option.textContent = `${person.displayName} (${email})`;
        approverSelect.appendChild(option);
      }
    });

    // Enable button only when a valid selection is made
    approverSelect.onchange = () => {
      confirmBtn.disabled = false;
    };

    modal.show();
  } catch (err) {
    console.error("Error loading people list:", err);
    approverSelect.innerHTML = '<option disabled selected>Unable to load contacts</option>';
  }
}

