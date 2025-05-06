document.addEventListener("DOMContentLoaded", () => {
  // --- DOM Elements ---
  // const htmlElement = document.documentElement; // No longer needed for theme
  const headerConnectionName = document.getElementById("headerConnectionName");
  const mobileSidebar = document.getElementById("mobileSidebar");
  const mobileSidebarBackdrop = document.getElementById(
    "mobileSidebarBackdrop",
  );
  const openMobileSidebarButton = document.getElementById(
    "openMobileSidebarButton",
  );
  const closeMobileSidebarButton = document.getElementById(
    "closeMobileSidebarButton",
  );

  const newConnectionButton = document.getElementById("newConnectionButton");
  const newConnectionButtonMobile = document.getElementById(
    "newConnectionButtonMobile",
  );
  const savedConnectionsListUI = document.getElementById(
    "savedConnectionsList",
  );
  const savedConnectionsListMobileUI = document.getElementById(
    "savedConnectionsListMobile",
  );

  const editorPane = document.getElementById("editorPane");
  const connectionFormTitle = document.getElementById("connectionFormTitle");
  const saveConnectionButton = document.getElementById("saveConnectionButton");

  // Import/Export
  const exportConnectionsButton = document.getElementById(
    "exportConnectionsButton",
  );
  const importConnectionsInput = document.getElementById(
    "importConnectionsInput",
  );

  // Tabs
  const editorTabsContainer = document.getElementById("editorTabs");
  const tabButtons = Array.from(
    editorTabsContainer.querySelectorAll(".tab-button"),
  );
  const tabContentContainer = document.getElementById("tabContentContainer");

  // General Tab Inputs
  const connectionNameInput = document.getElementById("connectionName");
  const apiUrlInput = document.getElementById("apiUrl");
  const usernameInput = document.getElementById("username");
  const passwordInput = document.getElementById("password");

  // Variables Tab
  const addVariableButton = document.getElementById("addVariableButton");
  const variablesContainer = document.getElementById("variablesContainer");
  const variablePlaceholder = variablesContainer.querySelector(
    ".variable-placeholder",
  );

  // OData Params Tab
  const paramFilterInput = document.getElementById("paramFilter");
  const paramOrderbyInput = document.getElementById("paramOrderby");
  const paramTopInput = document.getElementById("paramTop");
  const paramSkipInput = document.getElementById("paramSkip");
  const paramSelectInput = document.getElementById("paramSelect");

  // Actions
  const fetchButton = document.getElementById("fetchButton");
  const fetchAndDownloadButton = document.getElementById(
    "fetchAndDownloadButton",
  );
  const downloadButton = document.getElementById("downloadButton");

  // Activity Log
  const statusDiv = document.getElementById("status");
  const clearLogButton = document.getElementById("clearLogButton");

  // Modal
  const modal = document.getElementById("modal");
  const modalBackdrop = document.getElementById("modalBackdrop");
  const modalContent = document.getElementById("modalContent");
  const modalIconContainer = document.getElementById("modalIconContainer");
  const modalTitleElement = document.getElementById("modalTitle");
  const modalMessageElement = document.getElementById("modalMessage");
  const modalInputContainer = document.getElementById("modalInputContainer");
  const modalInputLabel = document.getElementById("modalInputLabel");
  const modalTextInput = document.getElementById("modalTextInput");
  const modalConfirmButton = document.getElementById("modalConfirmButton");
  const modalCancelButton = document.getElementById("modalCancelButton");

  // --- App State ---
  const STORAGE_KEY = "odataFetcherProXConnections_v3_light"; // Updated key
  let allFetchedData = [];
  let isFetching = false;
  let currentLoadedConnectionOriginalName = null;
  let currentModalConfirmCallback = null;

  // --- Icons (SVG strings) ---
  const ICONS = {
    trash: `<svg class="w-4 h-4" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M9 2a1 1 0 00-.894.553L7.382 4H4a1 1 0 000 2v10a2 2 0 002 2h8a2 2 0 002-2V6a1 1 0 100-2h-3.382l-.724-1.447A1 1 0 0011 2H9zM7 8a1 1 0 012 0v6a1 1 0 11-2 0V8zm5-1a1 1 0 00-1 1v6a1 1 0 102 0V8a1 1 0 00-1-1z" clip-rule="evenodd"></path></svg>`,
    warning: `<svg class="h-6 w-6 text-red-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126zM12 15.75h.007v.008H12v-.008z" /></svg>`,
    info: `<svg class="h-6 w-6 text-sky-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M11.25 11.25l.041-.02a.75.75 0 011.063.852l-.708 2.836a.75.75 0 001.063.853l.041-.021M21 12a9 9 0 11-18 0 9 9 0 0118 0zm-9-3.75h.008v.008H12V8.25z" /></svg>`,
    removeX: `<svg class="w-4 h-4" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clip-rule="evenodd"></path></svg>`,
    duplicate: `<svg class="w-4 h-4" fill="currentColor" viewBox="0 0 20 20"><path d="M7 9a2 2 0 012-2h6a2 2 0 012 2v6a2 2 0 01-2 2H9a2 2 0 01-2-2V9z"></path><path d="M3 3a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2V3a2 2 0 00-2-2H3zm2 2h10v10H5V5z"></path></svg>`,
  };

  // --- Utility Functions ---
  function updateStatus(message, type = "info") {
    const logEntry = document.createElement("div");
    // Simplified log colors for light mode
    let typeClasses = "text-slate-700"; // Default for info
    let timestampClasses = "text-sky-600";
    if (type === "error") typeClasses = "text-red-600";
    else if (type === "success") typeClasses = "text-green-600";
    else if (type === "warning") typeClasses = "text-yellow-600"; // Or orange-600

    logEntry.className = `mb-1 leading-normal ${typeClasses}`;
    const timestampSpan = document.createElement("span");
    timestampSpan.className = `${timestampClasses} mr-2`;
    timestampSpan.textContent = `[${new Date().toLocaleTimeString()}]`;
    const messageSpan = document.createElement("span");
    messageSpan.textContent = message;
    logEntry.appendChild(timestampSpan);
    logEntry.appendChild(messageSpan);
    statusDiv.appendChild(logEntry);
    statusDiv.scrollTop = statusDiv.scrollHeight;
  }

  // --- Modal Management ---
  function showModal({
    title,
    message,
    confirmText = "Confirm",
    cancelText = "Cancel",
    onConfirm,
    iconType = "info",
    input = null,
  }) {
    modalTitleElement.textContent = title;
    modalMessageElement.textContent = message;
    modalConfirmButton.textContent = confirmText;
    modalCancelButton.textContent = cancelText;

    modalIconContainer.innerHTML = ICONS[iconType] || ICONS.info;
    modalIconContainer.className =
      "mr-4 flex-shrink-0 flex items-center justify-center h-12 w-12 rounded-full sm:h-10 sm:w-10";

    modalConfirmButton.className =
      "w-full sm:w-auto inline-flex justify-center rounded-md border border-transparent shadow-sm px-4 py-2 text-base font-medium text-white focus:outline-none focus:ring-2 focus:ring-offset-2 sm:text-sm cursor-pointer";

    if (iconType === "warning" || iconType === "delete") {
      modalIconContainer.classList.add("bg-red-100");
      modalConfirmButton.classList.add(
        "bg-red-600",
        "hover:bg-red-700",
        "focus:ring-red-500",
      );
    } else {
      modalIconContainer.classList.add("bg-sky-100");
      modalConfirmButton.classList.add(
        "bg-sky-600",
        "hover:bg-sky-700",
        "focus:ring-sky-500",
      );
    }

    if (input) {
      modalInputContainer.classList.remove("hidden");
      modalInputLabel.textContent = input.label;
      modalTextInput.value = input.value || "";
      modalTextInput.placeholder = input.placeholder || "";
      setTimeout(() => modalTextInput.focus(), 50);
    } else {
      modalInputContainer.classList.add("hidden");
    }

    modal.classList.remove("hidden");
    setTimeout(() => {
      modalBackdrop.classList.remove("opacity-0");
      modalBackdrop.classList.add("opacity-100");
      modalContent.classList.remove("scale-95", "opacity-0");
      modalContent.classList.add("scale-100", "opacity-100");
    }, 10);
    currentModalConfirmCallback = onConfirm;
  }
  function hideModal() {
    modalBackdrop.classList.remove("opacity-100");
    modalBackdrop.classList.add("opacity-0");
    modalContent.classList.remove("scale-100", "opacity-100");
    modalContent.classList.add("scale-95", "opacity-0");
    setTimeout(() => {
      modal.classList.add("hidden");
      modalInputContainer.classList.add("hidden");
    }, 200);
    currentModalConfirmCallback = null;
  }
  modalConfirmButton.addEventListener("click", () => {
    if (currentModalConfirmCallback) {
      const inputValue = modalInputContainer.classList.contains("hidden")
        ? null
        : modalTextInput.value;
      currentModalConfirmCallback(inputValue);
    }
    hideModal();
  });
  modalCancelButton.addEventListener("click", hideModal);
  modalBackdrop.addEventListener("click", hideModal);

  // --- Sidebar Management ---
  function toggleMobileSidebar(open) {
    if (open) {
      mobileSidebar.classList.add("open");
      mobileSidebarBackdrop.classList.remove("hidden");
    } else {
      mobileSidebar.classList.remove("open");
      mobileSidebarBackdrop.classList.add("hidden");
    }
  }
  openMobileSidebarButton.addEventListener("click", () =>
    toggleMobileSidebar(true),
  );
  closeMobileSidebarButton.addEventListener("click", () =>
    toggleMobileSidebar(false),
  );
  mobileSidebarBackdrop.addEventListener("click", () =>
    toggleMobileSidebar(false),
  );

  // --- Tab Management ---
  tabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      setActiveTab(button.dataset.tab);
    });
  });

  function setActiveTab(tabIdToActivate) {
    tabButtons.forEach((btn) => {
      const currentTabId = btn.dataset.tab;
      const correspondingContent = document.getElementById(
        `tab-${currentTabId}`,
      );

      if (currentTabId === tabIdToActivate) {
        btn.classList.add("active"); // CSS in HTML handles the active look
        if (correspondingContent)
          correspondingContent.classList.remove("hidden");
      } else {
        btn.classList.remove("active");
        if (correspondingContent) correspondingContent.classList.add("hidden");
      }
    });
  }

  // --- Variable UI Management ---
  function createVariableRowDOM(name = "", value = "") {
    const row = document.createElement("div");
    row.className =
      "variable-row flex items-center space-x-2 bg-white p-2 rounded-md shadow-sm"; // Light mode only
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.placeholder = "Name (e.g., tenantId)";
    nameInput.className =
      "variable-name-input w-2/5 px-2.5 py-1.5 bg-transparent border border-slate-300 rounded-md text-sm focus:border-sky-500 focus:ring-1 focus:ring-sky-500 focus:outline-none transition-colors text-slate-700";
    nameInput.value = name;
    const valueInput = document.createElement("input");
    valueInput.type = "text";
    valueInput.placeholder = "Value";
    valueInput.className =
      "variable-value-input flex-grow px-2.5 py-1.5 bg-transparent border border-slate-300 rounded-md text-sm focus:border-sky-500 focus:ring-1 focus:ring-sky-500 focus:outline-none transition-colors text-slate-700";
    valueInput.value = value;
    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className =
      "remove-variable-button text-slate-400 hover:text-red-500 p-1.5 rounded-md cursor-pointer transition-colors duration-150";
    removeButton.innerHTML = ICONS.removeX;
    removeButton.onclick = () => {
      row.remove();
      checkVariablePlaceholder();
    };
    row.appendChild(nameInput);
    row.appendChild(valueInput);
    row.appendChild(removeButton);
    variablesContainer.appendChild(row);
    checkVariablePlaceholder();
  }
  function checkVariablePlaceholder() {
    if (variablesContainer.querySelector(".variable-row"))
      variablePlaceholder.classList.add("hidden");
    else variablePlaceholder.classList.remove("hidden");
  }
  addVariableButton.addEventListener("click", () => createVariableRowDOM());
  function getVariablesFromUI() {
    const variables = [];
    variablesContainer.querySelectorAll(".variable-row").forEach((row) => {
      const nameInput = row.querySelector(".variable-name-input");
      const valueInput = row.querySelector(".variable-value-input");
      if (nameInput && valueInput && nameInput.value.trim() !== "") {
        variables.push({
          name: nameInput.value.trim(),
          value: valueInput.value,
        });
      }
    });
    return variables;
  }
  function renderVariablesUI(varsArray = []) {
    variablesContainer.innerHTML = "";
    variablesContainer.appendChild(variablePlaceholder);
    if (varsArray && varsArray.length > 0)
      varsArray.forEach((v) => createVariableRowDOM(v.name, v.value));
    checkVariablePlaceholder();
  }

  // --- OData Parameters UI Management ---
  function getODataParamsFromUI() {
    return {
      filter: paramFilterInput.value.trim(),
      orderby: paramOrderbyInput.value.trim(),
      top: paramTopInput.value.trim(),
      skip: paramSkipInput.value.trim(),
      select: paramSelectInput.value.trim(),
    };
  }
  function renderODataParamsUI(params = {}) {
    paramFilterInput.value = params.filter || "";
    paramOrderbyInput.value = params.orderby || "";
    paramTopInput.value = params.top || "";
    paramSkipInput.value = params.skip || "";
    paramSelectInput.value = params.select || "";
  }

  // --- Variable Substitution ---
  function substituteVariables(textToProcess, variables) {
    let processedText = textToProcess;
    if (variables && variables.length > 0) {
      variables.forEach((variable) => {
        const placeholder = new RegExp(
          `\\{\\{${variable.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\}\\}`,
          "g",
        );
        processedText = processedText.replace(placeholder, variable.value);
      });
    }
    if (processedText.includes("://") || textToProcess !== processedText) {
      const unsubstituted = processedText.match(/\{\{([^}]+)\}\}/g);
      if (unsubstituted) {
        updateStatus(
          `Warning: Unsubstituted: ${unsubstituted.join(", ")} in "${processedText.substring(0, 70)}..."`,
          "warning",
        );
      }
    }
    return processedText;
  }

  // --- URL Construction with OData Params & Variable Substitution in Params ---
  function buildUrlWithODataParams(baseUrl, rawOdataParams, activeVariables) {
    const params = new URLSearchParams();
    const processParam = (paramValue) =>
      substituteVariables(paramValue, activeVariables);
    if (rawOdataParams.filter)
      params.append("$filter", processParam(rawOdataParams.filter));
    if (rawOdataParams.orderby)
      params.append("$orderby", processParam(rawOdataParams.orderby));
    if (rawOdataParams.top)
      params.append("$top", processParam(rawOdataParams.top));
    if (rawOdataParams.skip)
      params.append("$skip", processParam(rawOdataParams.skip));
    if (rawOdataParams.select)
      params.append("$select", processParam(rawOdataParams.select));
    const queryString = params.toString();
    return queryString
      ? `${baseUrl}${baseUrl.includes("?") ? "&" : "?"}${queryString}`
      : baseUrl;
  }

  // --- LocalStorage Connection Management ---
  function getSavedConnections() {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      return JSON.parse(stored).map((conn) => ({
        ...conn,
        variables: Array.isArray(conn.variables) ? conn.variables : [],
        odataParams:
          typeof conn.odataParams === "object" && conn.odataParams !== null
            ? conn.odataParams
            : {},
      }));
    }
    return [];
  }
  function saveConnectionsToStorage(connections) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(connections));
  }

  function createConnectionListItem(conn, listUI) {
    const li = document.createElement("li");
    li.className =
      "group flex justify-between items-center p-2.5 rounded-md hover:bg-slate-200 cursor-pointer transition-colors duration-150"; // Light mode hover
    li.dataset.connName = conn.name;

    const nameSpan = document.createElement("span");
    nameSpan.className =
      "text-sm text-slate-700 group-[.selected]:text-sky-600 group-[.selected]:font-semibold truncate mr-2"; // Light mode selected text
    nameSpan.textContent = conn.name;
    li.appendChild(nameSpan);

    const buttonsDiv = document.createElement("div");
    buttonsDiv.className =
      "flex items-center space-x-1 opacity-0 group-hover:opacity-100 group-[.selected]:opacity-100 transition-opacity duration-150";

    const duplicateBtn = document.createElement("button");
    duplicateBtn.className = "text-slate-400 hover:text-sky-500 p-1 rounded-md"; // Light mode icon colors
    duplicateBtn.innerHTML = ICONS.duplicate;
    duplicateBtn.title = `Duplicate ${conn.name}`;
    duplicateBtn.onclick = (e) => {
      e.stopPropagation();
      handleDuplicateConnection(conn.name);
    };
    buttonsDiv.appendChild(duplicateBtn);

    const deleteBtn = document.createElement("button");
    deleteBtn.className = "text-slate-400 hover:text-red-500 p-1 rounded-md"; // Light mode icon colors
    deleteBtn.innerHTML = ICONS.trash;
    deleteBtn.title = `Delete ${conn.name}`;
    deleteBtn.onclick = (e) => {
      e.stopPropagation();
      showModal({
        title: "Delete Connection",
        message: `Are you sure you want to delete "${conn.name}"? This action cannot be undone.`,
        confirmText: "Delete",
        iconType: "delete",
        onConfirm: () => handleDeleteConnection(conn.name),
      });
    };
    buttonsDiv.appendChild(deleteBtn);
    li.appendChild(buttonsDiv);

    li.onclick = () => {
      handleLoadConnection(conn.name);
      if (listUI === savedConnectionsListMobileUI) toggleMobileSidebar(false);
    };
    return li;
  }

  function renderSavedConnectionsList() {
    const connections = getSavedConnections();
    [savedConnectionsListUI, savedConnectionsListMobileUI].forEach(
      (list) => (list.innerHTML = ""),
    );
    const placeholderLi = (text) => {
      const li = document.createElement("li");
      li.textContent = text;
      li.className = "text-center text-slate-400 text-sm py-4";
      return li;
    };
    if (connections.length === 0) {
      savedConnectionsListUI.appendChild(placeholderLi("No connections yet."));
      savedConnectionsListMobileUI.appendChild(
        placeholderLi("No connections yet."),
      );
      return;
    }
    connections.sort((a, b) => a.name.localeCompare(b.name));
    connections.forEach((conn) => {
      savedConnectionsListUI.appendChild(
        createConnectionListItem(conn, savedConnectionsListUI),
      );
      savedConnectionsListMobileUI.appendChild(
        createConnectionListItem(conn, savedConnectionsListMobileUI),
      );
    });
    updateSelectedListItemUI(currentLoadedConnectionOriginalName);
  }

  function clearForm(showStatus = true) {
    connectionNameInput.value = "";
    apiUrlInput.value = "";
    usernameInput.value = "";
    passwordInput.value = "";
    renderVariablesUI([]);
    renderODataParamsUI({});
    currentLoadedConnectionOriginalName = null;
    connectionFormTitle.textContent = "New Connection";
    headerConnectionName.textContent = "New Connection";
    updateSelectedListItemUI(null);
    setActiveTab("general");
    if (showStatus)
      updateStatus("Form cleared. Enter details for a new connection.");
    editorPane.scrollTop = 0;
  }

  function updateSelectedListItemUI(selectedName) {
    [savedConnectionsListUI, savedConnectionsListMobileUI].forEach((list) => {
      Array.from(list.children).forEach((item) => {
        if (item.dataset && item.dataset.connName === selectedName) {
          item.classList.add("selected", "bg-sky-100"); // Light mode selected bg
        } else {
          item.classList.remove("selected", "bg-sky-100");
        }
      });
    });
  }

  [newConnectionButton, newConnectionButtonMobile].forEach((btn) => {
    btn.addEventListener("click", () => {
      clearForm();
      connectionNameInput.focus();
      if (btn === newConnectionButtonMobile) toggleMobileSidebar(false);
    });
  });

  function handleLoadConnection(name, isDuplicate = false) {
    const connections = getSavedConnections();
    const conn = connections.find((c) => c.name === name);
    if (conn) {
      connectionNameInput.value = isDuplicate
        ? `Copy of ${conn.name}`
        : conn.name;
      apiUrlInput.value = conn.url;
      usernameInput.value = conn.username || "";
      passwordInput.value = conn.password || "";
      renderVariablesUI(conn.variables || []);
      renderODataParamsUI(conn.odataParams || {});
      currentLoadedConnectionOriginalName = isDuplicate ? null : conn.name;
      const title = isDuplicate
        ? `New Connection (Copy)`
        : `Editing: ${conn.name}`;
      connectionFormTitle.textContent = title;
      headerConnectionName.textContent = connectionNameInput.value;
      updateStatus(
        isDuplicate
          ? `Duplicated "${conn.name}". Ready to save as new.`
          : `Loaded connection: "${conn.name}".`,
      );
      updateSelectedListItemUI(isDuplicate ? null : conn.name);
      setActiveTab("general");
      editorPane.scrollTop = 0;
      if (isDuplicate) connectionNameInput.focus();
    }
  }

  function handleDuplicateConnection(name) {
    handleLoadConnection(name, true);
    if (document.documentElement.clientWidth < 768) toggleMobileSidebar(false);
  }

  saveConnectionButton.addEventListener("click", () => {
    const nameFromForm = connectionNameInput.value.trim();
    const urlFromForm = apiUrlInput.value.trim();
    const userFromForm = usernameInput.value.trim();
    const passFromForm = passwordInput.value;
    const variablesFromUI = getVariablesFromUI();
    const odataParamsFromUI = getODataParamsFromUI();

    if (!nameFromForm) {
      showModal({
        title: "Validation Error",
        message: "Connection Name is required.",
        iconType: "warning",
        confirmText: "OK",
      });
      connectionNameInput.focus();
      return;
    }
    if (!urlFromForm) {
      showModal({
        title: "Validation Error",
        message: "API URL is required.",
        iconType: "warning",
        confirmText: "OK",
      });
      apiUrlInput.focus();
      return;
    }

    let connections = getSavedConnections();
    const newOrUpdatedConnData = {
      name: nameFromForm,
      url: urlFromForm,
      username: userFromForm,
      password: passFromForm,
      variables: variablesFromUI,
      odataParams: odataParamsFromUI,
    };

    const existingByNameIndex = connections.findIndex(
      (c) => c.name === nameFromForm,
    );
    let actionType = "new";

    if (currentLoadedConnectionOriginalName) {
      if (currentLoadedConnectionOriginalName === nameFromForm)
        actionType = "update";
      else actionType = "rename";
    }

    if (existingByNameIndex !== -1) {
      if (
        actionType === "update" &&
        connections[existingByNameIndex].name ===
          currentLoadedConnectionOriginalName
      ) {
        connections[existingByNameIndex] = newOrUpdatedConnData;
        finishSave(
          connections,
          newOrUpdatedConnData.name,
          `Connection "${nameFromForm}" updated.`,
        );
      } else {
        showModal({
          title: "Name Conflict",
          message: `A connection named "${nameFromForm}" already exists. Overwrite it?`,
          confirmText: "Overwrite",
          iconType: "warning",
          onConfirm: () => {
            if (actionType === "rename") {
              connections = connections.filter(
                (c) => c.name !== currentLoadedConnectionOriginalName,
              );
            }
            const finalExistingIndex = connections.findIndex(
              (c) => c.name === nameFromForm,
            );
            if (finalExistingIndex !== -1)
              connections[finalExistingIndex] = newOrUpdatedConnData;
            else connections.push(newOrUpdatedConnData);

            let msg = `Connection "${nameFromForm}" saved.`;
            if (actionType === "rename")
              msg = `Renamed "${currentLoadedConnectionOriginalName}" to "${nameFromForm}" and saved.`;
            else msg = `Connection "${nameFromForm}" overwritten.`;
            finishSave(connections, newOrUpdatedConnData.name, msg);
          },
        });
      }
    } else {
      if (actionType === "rename") {
        connections = connections.filter(
          (c) => c.name !== currentLoadedConnectionOriginalName,
        );
      }
      connections.push(newOrUpdatedConnData);
      let msg = `New connection "${nameFromForm}" saved.`;
      if (actionType === "rename")
        msg = `Renamed "${currentLoadedConnectionOriginalName}" to "${nameFromForm}" and saved.`;
      finishSave(connections, newOrUpdatedConnData.name, msg);
    }
  });

  function finishSave(connections, newName, message) {
    saveConnectionsToStorage(connections);
    currentLoadedConnectionOriginalName = newName;
    const title = `Editing: ${newName}`;
    connectionFormTitle.textContent = title;
    headerConnectionName.textContent = newName;
    renderSavedConnectionsList();
    updateStatus(message, "success");
  }

  function handleDeleteConnection(nameToDelete) {
    let connections = getSavedConnections();
    connections = connections.filter((c) => c.name !== nameToDelete);
    saveConnectionsToStorage(connections);
    if (currentLoadedConnectionOriginalName === nameToDelete) clearForm(false);
    renderSavedConnectionsList();
    updateStatus(`Connection "${nameToDelete}" deleted.`, "success");
  }

  // --- Import/Export Connections ---
  exportConnectionsButton.addEventListener("click", () => {
    const connections = getSavedConnections();
    if (connections.length === 0) {
      showModal({
        title: "Export Failed",
        message: "No connections to export.",
        iconType: "info",
        confirmText: "OK",
      });
      return;
    }
    const filename = `odata_fetcher_connections_${new Date().toISOString().slice(0, 10)}.json`;
    const jsonStr = JSON.stringify(connections, null, 2);
    const blob = new Blob([jsonStr], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    updateStatus("Connections exported successfully.", "success");
  });

  importConnectionsInput.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const importedConnections = JSON.parse(e.target.result);
        if (!Array.isArray(importedConnections)) {
          throw new Error("Imported file is not a valid connections array.");
        }
        const validConnections = importedConnections.filter(
          (conn) => conn.name && conn.url,
        );

        if (validConnections.length === 0 && importedConnections.length > 0) {
          showModal({
            title: "Import Warning",
            message:
              "No valid connections found in the file (missing name or url).",
            iconType: "warning",
            confirmText: "OK",
          });
          return;
        }

        let currentConnections = getSavedConnections();
        let importedCount = 0;
        let updatedCount = 0;

        validConnections.forEach((importedConn) => {
          const existingIndex = currentConnections.findIndex(
            (c) => c.name === importedConn.name,
          );
          const fullImportedConn = {
            name: importedConn.name,
            url: importedConn.url,
            username: importedConn.username || "",
            password: importedConn.password || "",
            variables: Array.isArray(importedConn.variables)
              ? importedConn.variables
              : [],
            odataParams:
              typeof importedConn.odataParams === "object" &&
              importedConn.odataParams !== null
                ? importedConn.odataParams
                : {},
          };

          if (existingIndex !== -1) {
            currentConnections[existingIndex] = fullImportedConn;
            updatedCount++;
          } else {
            currentConnections.push(fullImportedConn);
            importedCount++;
          }
        });
        saveConnectionsToStorage(currentConnections);
        renderSavedConnectionsList();
        clearForm(false);
        updateStatus(
          `${importedCount} new connections imported, ${updatedCount} connections updated. Please select one.`,
          "success",
        );
      } catch (error) {
        console.error("Import error:", error);
        showModal({
          title: "Import Error",
          message: `Failed to import connections: ${error.message}`,
          iconType: "error",
          confirmText: "OK",
        });
      } finally {
        importConnectionsInput.value = "";
      }
    };
    reader.readAsText(file);
  });

  // --- OData Fetching Logic ---
  async function fetchOData(rawUrlWithVars, username, password, page = 1) {
    let currentUrl;
    const activeVariables = getVariablesFromUI();
    if (page === 1) {
      const substitutedBaseUrl = substituteVariables(
        rawUrlWithVars,
        activeVariables,
      );
      const rawODataParams = getODataParamsFromUI();
      currentUrl = buildUrlWithODataParams(
        substitutedBaseUrl,
        rawODataParams,
        activeVariables,
      );
      updateStatus(
        `Effective URL (Page 1): ${currentUrl.length > 120 ? currentUrl.substring(0, 120) + "..." : currentUrl}`,
        "info",
      );
    } else {
      currentUrl = rawUrlWithVars;
    }
    updateStatus(
      `Fetching page ${page}: ${currentUrl.length > 120 ? currentUrl.substring(0, 120) + "..." : currentUrl}`,
    );
    fetchButton.disabled = true;
    fetchAndDownloadButton.disabled = true;
    downloadButton.disabled = true;

    const headers = new Headers();
    if (username || password)
      headers.append(
        "Authorization",
        "Basic " + btoa(username + ":" + password),
      );
    headers.append("Accept", "application/json;odata.metadata=minimal");
    try {
      const response = await fetch(currentUrl, {
        method: "GET",
        headers: headers,
      });
      if (!response.ok) {
        let errorMsg = `HTTP error! Status: ${response.status} on URL: ${currentUrl.substring(0, 100)}...`;
        try {
          const errorData = await response.json();
          errorMsg += ` - ${errorData.error?.message || JSON.stringify(errorData)}`;
        } catch (e) {
          errorMsg += ` - ${response.statusText}`;
        }
        throw new Error(errorMsg);
      }
      const data = await response.json();
      if (data.value && Array.isArray(data.value)) {
        allFetchedData.push(...data.value);
        updateStatus(
          `Fetched ${data.value.length} records from page ${page}. Total: ${allFetchedData.length}.`,
        );
      } else {
        console.warn("Response did not contain a 'value' array:", data);
        if (page === 1 && allFetchedData.length === 0) {
          updateStatus(
            `Warning: Page ${page} response did not contain 'value' array.`,
            "warning",
          );
        }
      }
      const nextLink = data["@odata.nextLink"];
      if (nextLink) {
        let resolvedNextLink;
        try {
          const initialApiUrlForBase = new URL(apiUrlInput.value.trim());
          resolvedNextLink = new URL(
            nextLink,
            initialApiUrlForBase.origin +
              initialApiUrlForBase.pathname.substring(
                0,
                initialApiUrlForBase.pathname.lastIndexOf("/") + 1,
              ),
          ).href;
        } catch (e) {
          resolvedNextLink = nextLink;
        }
        await fetchOData(resolvedNextLink, username, password, page + 1);
      }
    } catch (error) {
      throw error;
    }
  }

  // --- CSV Generation and Download ---
  function escapeCsvCell(cellData) {
    if (cellData == null) return "";
    const stringData = String(cellData);
    if (
      stringData.includes(",") ||
      stringData.includes("\n") ||
      stringData.includes('"')
    ) {
      return `"${stringData.replace(/"/g, '""')}"`;
    }
    return stringData;
  }
  function generateAndDownloadCSV(filenameSuffix = "") {
    if (allFetchedData.length === 0) {
      updateStatus("No data to download.", "error");
      return false;
    }
    updateStatus("Generating CSV...", "info");
    const headers =
      allFetchedData.length > 0 ? Object.keys(allFetchedData[0]) : [];
    if (headers.length === 0) {
      updateStatus(
        "Cannot generate CSV: Data is empty or has no properties.",
        "error",
      );
      return false;
    }
    let csvContent = headers.map(escapeCsvCell).join(",") + "\r\n";
    allFetchedData.forEach((rowObject) => {
      const row = headers.map((header) => escapeCsvCell(rowObject[header]));
      csvContent += row.join(",") + "\r\n";
    });
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    let filename = `odata_export${filenameSuffix}.csv`;
    const connName = connectionNameInput.value.trim();
    if (connName) {
      filename = `${connName.replace(/[^a-z0-9_.-]/gi, "_")}_export${filenameSuffix}.csv`;
    } else {
      try {
        const activeVariables = getVariablesFromUI();
        const substitutedBaseUrl = substituteVariables(
          apiUrlInput.value.trim(),
          activeVariables,
        );
        const urlPath = new URL(substitutedBaseUrl).pathname;
        const lastSegment =
          urlPath.substring(urlPath.lastIndexOf("/") + 1) || "data";
        filename = `${lastSegment.replace(/[^a-z0-9_.-]/gi, "_")}_export${filenameSuffix}.csv`;
      } catch (e) {
        /* ignore */
      }
    }
    if (link.download !== undefined) {
      const url = URL.createObjectURL(blob);
      link.setAttribute("href", url);
      link.setAttribute("download", filename);
      link.style.visibility = "hidden";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      updateStatus(
        `CSV "${filename}" generated and download initiated.`,
        "success",
      );
      return true;
    } else {
      updateStatus("CSV download not supported by your browser.", "error");
      return false;
    }
  }

  // --- Main Action Event Listeners ---
  async function performFetch(andDownload = false) {
    if (isFetching) return;
    const rawUrlFromInput = apiUrlInput.value.trim();
    const usernameVal = usernameInput.value.trim();
    const passwordVal = passwordInput.value;

    if (!rawUrlFromInput) {
      showModal({
        title: "Missing URL",
        message: "API URL is required to fetch data.",
        iconType: "warning",
        confirmText: "OK",
      });
      apiUrlInput.focus();
      return;
    }
    allFetchedData = [];
    isFetching = true;
    fetchButton.disabled = true;
    fetchAndDownloadButton.disabled = true;
    downloadButton.disabled = true;
    updateStatus("Starting data fetch...", "info");
    try {
      await fetchOData(rawUrlFromInput, usernameVal, passwordVal);
      if (allFetchedData.length > 0) {
        updateStatus(
          `Successfully fetched ${allFetchedData.length} records.`,
          "success",
        );
        if (andDownload) {
          generateAndDownloadCSV(
            `_${new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-")}`,
          );
        }
      } else {
        updateStatus(
          "Fetch complete. No records found or 'value' array was missing.",
          "info",
        );
      }
    } catch (error) {
      updateStatus(`Fetch Error: ${error.message}`, "error");
      console.error("Fetch error:", error);
    } finally {
      isFetching = false;
      fetchButton.disabled = false;
      fetchAndDownloadButton.disabled = false;
      if (allFetchedData.length > 0) {
        downloadButton.disabled = false;
      }
    }
  }

  fetchButton.addEventListener("click", () => performFetch(false));
  fetchAndDownloadButton.addEventListener("click", () => performFetch(true));
  downloadButton.addEventListener("click", () =>
    generateAndDownloadCSV(
      `_${new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-")}`,
    ),
  );

  clearLogButton.addEventListener("click", () => {
    statusDiv.innerHTML = "";
  });

  // --- Initialization ---
  function init() {
    // loadTheme(); // Removed theme loading
    renderSavedConnectionsList();
    clearLogButton.click();
    updateStatus(
      "Welcome! Create a new connection or load an existing one to begin.",
    );
    clearForm(false);
    setActiveTab("general");
    checkVariablePlaceholder();
  }
  init();
});
