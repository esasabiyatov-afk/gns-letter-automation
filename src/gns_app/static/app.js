(() => {
  const fileInput = document.querySelector("#pdf-file");
  const dropZone = document.querySelector(".drop-zone");
  if (fileInput && dropZone) {
    fileInput.addEventListener("change", () => {
      const file = fileInput.files?.[0];
      if (file) {
        dropZone.querySelector("strong").textContent = file.name;
        dropZone.classList.add("has-file");
      }
    });
  }

  document.querySelectorAll("[data-open-dialog]").forEach((button) => {
    button.addEventListener("click", () => {
      document.getElementById(button.dataset.openDialog)?.showModal();
    });
  });
  document.querySelectorAll("[data-close-dialog]").forEach((button) => {
    button.addEventListener("click", () => button.closest("dialog")?.close());
  });

  document.querySelectorAll("[data-odb-result]").forEach((button) => {
    button.addEventListener("click", async () => {
      const row = button.closest("[data-odb-row]");
      const odbCase = button.closest("[data-odb-case]");
      const inn = row?.querySelector("input[name='taxpayer_inn']")?.value;
      if (!row || !odbCase || !inn || row.dataset.saving === "true") return;

      row.dataset.saving = "true";
      row.querySelectorAll("button").forEach((item) => { item.disabled = true; });
      const body = new FormData();
      body.append("taxpayer_inn", inn);
      body.append("odb_result", button.dataset.odbResult);
      try {
        const response = await fetch(row.dataset.odbUrl, {
          method: "POST",
          body,
        });
        const payload = await response.json();
        if (!response.ok || !payload.ok) {
          throw new Error(payload.error || "Не удалось сохранить проверку ОДБ");
        }
        row.classList.add("is-removing");
        window.setTimeout(() => {
          row.remove();
          if (!odbCase.querySelector("[data-odb-row]")) odbCase.remove();
          const panel = document.querySelector(".odb-today-panel");
          if (panel && !panel.querySelector("[data-odb-case]")) panel.remove();
        }, 180);
      } catch (error) {
        row.dataset.saving = "false";
        row.querySelectorAll("button").forEach((item) => { item.disabled = false; });
        window.alert(error.message);
      }
    });
  });

  const viewerTabs = document.querySelectorAll("[data-viewer-tab]");
  const viewerPanels = document.querySelectorAll("[data-viewer-panel]");
  if (viewerTabs.length && viewerPanels.length) {
    viewerTabs.forEach((button) => {
      button.addEventListener("click", () => {
        const selected = button.dataset.viewerTab;
        viewerTabs.forEach((tab) => {
          tab.classList.toggle("active", tab === button);
        });
        viewerPanels.forEach((panel) => {
          panel.hidden = panel.dataset.viewerPanel !== selected;
        });
      });
    });
  }

  document.querySelector("#copy-ocr")?.addEventListener("click", async (event) => {
    const text = document.querySelector("#ocr-selectable")?.textContent || "";
    try {
      await navigator.clipboard.writeText(text);
      event.currentTarget.textContent = "Скопировано";
      window.setTimeout(() => { event.currentTarget.textContent = "Копировать текст"; }, 1600);
    } catch (_) {
      const selection = window.getSelection();
      const range = document.createRange();
      range.selectNodeContents(document.querySelector("#ocr-selectable"));
      selection.removeAllRanges();
      selection.addRange(range);
    }
  });

  const autoGrow = (element) => {
    if (!element) return;
    element.style.height = "auto";
    element.style.height = `${Math.max(element.scrollHeight, 44)}px`;
  };
  const bindAutoGrow = (root = document) => {
    root.querySelectorAll?.("textarea[data-autogrow]").forEach((textarea) => {
      if (textarea.dataset.autogrowBound) return;
      textarea.dataset.autogrowBound = "1";
      textarea.addEventListener("input", () => autoGrow(textarea));
      autoGrow(textarea);
    });
  };
  bindAutoGrow();

  const taxpayerList = document.querySelector("#taxpayer-list");
  const addTaxpayer = document.querySelector("#add-taxpayer");
  if (taxpayerList) {
    const showRegistryResult = (row, data) => {
      const result = row.querySelector(".registry-live-result");
      if (!result) return;
      result.replaceChildren();
      result.className = "registry-live-result";
      if (data.status === "not_applicable") return;
      const message = document.createElement("span");
      message.textContent = data.message || "Не удалось выполнить сверку.";
      result.appendChild(message);
      if (data.status === "found" && data.official_name) {
        result.classList.add("is-found");
        message.textContent = `ОсОО.KG: ${data.official_name}`;
        const useName = document.createElement("button");
        useName.type = "button";
        useName.className = "text-button registry-name-choice";
        useName.dataset.registryName = data.official_name;
        useName.textContent = "Подставить название";
        result.appendChild(useName);
      } else if (["error", "multiple", "classification_uncertain"].includes(data.status)) {
        result.classList.add("is-warning");
      }
    };

    const scheduleRegistryLookup = (row) => {
      window.clearTimeout(row._registryTimer);
      row._registryController?.abort();
      const innInput = row.querySelector("input[name='taxpayer_inn']");
      const nameInput = row.querySelector("[name='taxpayer_name']");
      const result = row.querySelector(".registry-live-result");
      const inn = (innInput?.value || "").replace(/\D/g, "");
      if (inn.length !== 14) {
        row.dataset.registryLookupKey = "";
        result?.replaceChildren();
        return;
      }
      const lookupKey = `${inn}|${inn.startsWith("4") ? nameInput?.value || "" : ""}`;
      if (row.dataset.registryLookupKey === lookupKey) return;
      row.dataset.registryLookupKey = lookupKey;
      if (result) {
        result.className = "registry-live-result";
        result.textContent = "Проверяем ИНН в ОсОО.KG…";
      }
      row._registryTimer = window.setTimeout(async () => {
        const controller = new AbortController();
        row._registryController = controller;
        const params = new URLSearchParams({
          inn,
          name: nameInput?.value || "",
        });
        try {
          const response = await fetch(`/api/registry/suggestion?${params}`, {
            signal: controller.signal,
            headers: { Accept: "application/json" },
          });
          if (!response.ok) throw new Error("local_api_error");
          if (row.dataset.registryLookupKey === lookupKey) {
            showRegistryResult(row, await response.json());
          }
        } catch (error) {
          if (error.name !== "AbortError" && row.dataset.registryLookupKey === lookupKey) {
            showRegistryResult(row, {
              status: "error",
              message: "Не удалось выполнить сверку. Повторите ввод ИНН.",
            });
          }
        }
      }, 450);
    };

    const bindRegistryLookup = (root) => {
      root.querySelectorAll(".taxpayer-row").forEach((row) => {
        row.querySelector("input[name='taxpayer_inn']")?.addEventListener(
          "input", () => scheduleRegistryLookup(row)
        );
        row.querySelector("[name='taxpayer_name']")?.addEventListener(
          "input", () => {
            if ((row.querySelector("input[name='taxpayer_inn']")?.value || "").startsWith("4")) {
              scheduleRegistryLookup(row);
            }
          }
        );
        scheduleRegistryLookup(row);
      });
    };

    const bindRemove = (root) => {
      root.querySelectorAll(".remove-row").forEach((button) => {
        button.onclick = () => {
          const rows = taxpayerList.querySelectorAll(".taxpayer-row");
          if (rows.length > 1) button.closest(".taxpayer-row")?.remove();
          else {
            rows[0].querySelectorAll("input").forEach((input) => { input.value = ""; });
          }
        };
      });
    };
    addTaxpayer?.addEventListener("click", () => {
      const source = taxpayerList.querySelector(".taxpayer-row");
      if (!source) return;
      const row = source.cloneNode(true);
      row.querySelectorAll("input, textarea").forEach((input) => { input.value = ""; });
      row.querySelector(".registry-live-result")?.replaceChildren();
      row.querySelectorAll(".registry-state, .registry-name-choice").forEach(
        (element) => element.remove()
      );
      row.dataset.registryLookupKey = "";
      taxpayerList.appendChild(row);
      bindRemove(row);
      bindRegistryLookup(row);
      bindAutoGrow(row);
    });
    bindRemove(taxpayerList);
    bindRegistryLookup(taxpayerList);

    taxpayerList.addEventListener("click", (event) => {
      const button = event.target.closest("[data-registry-name]");
      if (!button) return;
      const input = button.closest(".taxpayer-row")?.querySelector(
        "[name='taxpayer_name']"
      );
      if (input) {
        input.value = button.dataset.registryName || "";
        autoGrow(input);
        button.closest(".registry-live-result")?.replaceChildren();
        button.closest(".taxpayer-name-field")?.querySelectorAll(
          ".registry-state, small, .registry-name-choice"
        ).forEach((element) => element.remove());
        input.focus();
      }
    });
  }

  document.querySelectorAll("[data-office-value]").forEach((button) => {
    button.addEventListener("click", () => {
      const input = document.querySelector("[name='district_place']");
      if (input) {
        input.value = button.dataset.officeValue || "";
        autoGrow(input);
        input.focus();
      }
    });
  });

  const officeInput = document.querySelector("[name='district_place']");
  const officePicker = document.querySelector("[data-office-picker]");
  if (officeInput && officePicker) {
    const officeButtons = [...officePicker.querySelectorAll("[data-office-value]")];
    const updateOfficePicker = () => {
      const query = officeInput.value.trim().toLocaleLowerCase("ru");
      let visible = 0;
      officeButtons.forEach((button) => {
        const matches = query.length >= 2 && (button.dataset.officeSearch || "")
          .toLocaleLowerCase("ru").includes(query);
        const show = matches && visible < 6;
        button.hidden = !show;
        if (show) visible += 1;
      });
      officePicker.hidden = visible === 0;
    };
    officeInput.addEventListener("input", updateOfficePicker);
    officeInput.addEventListener("focus", updateOfficePicker);
    officePicker.addEventListener("mousedown", (event) => event.preventDefault());
    officeButtons.forEach((button) => {
      button.addEventListener("click", () => {
        officeInput.value = button.dataset.officeValue || "";
        autoGrow(officeInput);
        officePicker.hidden = true;
        officeInput.focus();
      });
    });
    officeInput.addEventListener("blur", () => {
      window.setTimeout(() => { officePicker.hidden = true; }, 100);
    });
  }

  const positionChoice = document.querySelector("#recipient-position-choice");
  const positionValue = document.querySelector("#recipient-position-value");
  const positionCustom = document.querySelector("#recipient-position-custom");
  if (positionChoice && positionValue && positionCustom) {
    const syncPosition = () => {
      const custom = positionChoice.value === "other";
      positionCustom.hidden = !custom;
      positionValue.value = custom ? positionCustom.value.trim() : positionChoice.value;
      if (custom) autoGrow(positionCustom);
    };
    positionChoice.addEventListener("change", syncPosition);
    positionCustom.addEventListener("input", syncPosition);
    syncPosition();
  }

  const recipientFullName = document.querySelector("[name='recipient_full_name']");
  const recipientDisplay = document.querySelector("[data-recipient-display]");
  if (recipientFullName && recipientDisplay) {
    let displayTimer = null;
    const updateRecipientDisplay = () => {
      if (recipientDisplay.dataset.manualEdit === "1") return;
      window.clearTimeout(displayTimer);
      displayTimer = window.setTimeout(async () => {
        const params = new URLSearchParams({ full_name: recipientFullName.value });
        const response = await fetch(`/api/recipient-display?${params}`, {
          headers: { Accept: "application/json" },
        });
        if (response.ok && recipientDisplay.dataset.manualEdit !== "1") {
          recipientDisplay.value = (await response.json()).display_name || "";
        }
      }, 220);
    };
    recipientDisplay.addEventListener("dblclick", () => {
      recipientDisplay.readOnly = false;
      recipientDisplay.dataset.manualEdit = "1";
      recipientDisplay.focus();
      recipientDisplay.select();
    });
    recipientFullName.addEventListener("input", updateRecipientDisplay);
    updateRecipientDisplay();
  }

  const recipientPicker = document.querySelector("[data-recipient-picker]");
  if (recipientFullName && recipientPicker) {
    let recipientTimer = null;
    const hideRecipientPicker = () => { recipientPicker.hidden = true; };
    const chooseRecipient = (item) => {
      recipientFullName.value = item.full_name || "";
      autoGrow(recipientFullName);
      if (recipientDisplay) {
        recipientDisplay.value = item.display_name || "";
        recipientDisplay.readOnly = true;
        recipientDisplay.dataset.manualEdit = "0";
      }
      if (officeInput && item.district_place) {
        officeInput.value = item.district_place;
        autoGrow(officeInput);
      }
      if (positionChoice && positionValue && item.position) {
        const preset = [...positionChoice.options].some(
          (option) => option.value === item.position
        );
        positionChoice.value = preset ? item.position : "other";
        positionCustom.value = preset ? "" : item.position;
        positionChoice.dispatchEvent(new Event("change"));
      }
      hideRecipientPicker();
      recipientFullName.focus();
    };
    const updateRecipientPicker = () => {
      window.clearTimeout(recipientTimer);
      const query = recipientFullName.value.trim();
      if (query.length < 2) {
        hideRecipientPicker();
        return;
      }
      recipientTimer = window.setTimeout(async () => {
        const params = new URLSearchParams({ query });
        const response = await fetch(`/api/recipient-suggestions?${params}`, {
          headers: { Accept: "application/json" },
        });
        if (!response.ok) return;
        const items = (await response.json()).items || [];
        recipientPicker.replaceChildren();
        items.forEach((item) => {
          const button = document.createElement("button");
          button.type = "button";
          const strong = document.createElement("strong");
          strong.textContent = item.full_name || "";
          const small = document.createElement("small");
          small.textContent = item.district_place || item.position || "";
          button.append(strong, small);
          button.addEventListener("mousedown", (event) => event.preventDefault());
          button.addEventListener("click", () => chooseRecipient(item));
          recipientPicker.append(button);
        });
        recipientPicker.hidden = items.length === 0;
      }, 180);
    };
    recipientFullName.addEventListener("input", updateRecipientPicker);
    recipientFullName.addEventListener("focus", updateRecipientPicker);
    recipientFullName.addEventListener("blur", () => {
      window.setTimeout(hideRecipientPicker, 120);
    });
  }

  const periodSummary = document.querySelector("[data-period-summary]");
  const periodStart = document.querySelector("[name='period_start']");
  const periodEnd = document.querySelector("[name='period_end']");
  const periodRoutes = document.querySelectorAll("[name='period_route']");
  if (periodSummary && periodStart && periodEnd) {
    const updatePeriodSummary = () => {
      const start = periodStart.value;
      const end = periodEnd.value;
      const route = document.querySelector("[name='period_route']:checked")?.value;
      periodSummary.classList.remove("requires-odb", "no-odb", "period-missing");
      if ((!start || !end) && route === "odb") {
        periodSummary.classList.add("requires-odb");
        periodSummary.innerHTML = "<strong>Требуется проверка в ОДБ</strong><small>Выбрано сотрудником, так как точные даты не читаются.</small>";
      } else if ((!start || !end) && route === "no_odb") {
        periodSummary.classList.add("no-odb");
        periodSummary.innerHTML = "<strong>Проверка в ОДБ не требуется</strong><small>Выбрано сотрудником, так как точные даты не читаются.</small>";
      } else if (!start || !end) {
        periodSummary.classList.add("period-missing");
        periodSummary.innerHTML = "<strong>Период нужно уточнить</strong><small>Он определяет, требуется ли проверка в ОДБ.</small>";
      } else if (start < periodSummary.dataset.periodThreshold) {
        periodSummary.classList.add("requires-odb");
        periodSummary.innerHTML = `<strong>Требуется проверка в ОДБ</strong><small>Начало периода раньше ${periodSummary.dataset.periodThreshold}.</small>`;
      } else {
        periodSummary.classList.add("no-odb");
        periodSummary.innerHTML = `<strong>Проверка в ОДБ не требуется</strong><small>Период начинается с ${start}.</small>`;
      }
    };
    periodStart.addEventListener("change", updatePeriodSummary);
    periodEnd.addEventListener("change", updatePeriodSummary);
    periodRoutes.forEach((radio) => radio.addEventListener("change", updatePeriodSummary));
    updatePeriodSummary();
  }

  let activeTextInput = null;
  document.querySelectorAll("input[type='text'], input:not([type]), textarea").forEach((input) => {
    input.addEventListener("focus", () => { activeTextInput = input; });
  });
  document.querySelectorAll("[data-letter]").forEach((button) => {
    button.addEventListener("click", () => {
      if (!activeTextInput) return;
      const start = activeTextInput.selectionStart ?? activeTextInput.value.length;
      const end = activeTextInput.selectionEnd ?? start;
      activeTextInput.setRangeText(button.dataset.letter, start, end, "end");
      activeTextInput.focus();
    });
  });

  const letterFields = document.querySelector("#letter-fields");
  const typeRadios = document.querySelectorAll("input[name='page_type']");
  if (letterFields && typeRadios.length) {
    const verification = document.querySelector("[data-letter-verification]");
    const verificationWrapper = document.querySelector(
      "[data-letter-verification-wrapper]"
    );
    const update = () => {
      const selected = document.querySelector("input[name='page_type']:checked")?.value;
      letterFields.classList.toggle("is-disabled", selected !== "letter");
      letterFields.querySelectorAll("input, textarea, select").forEach((input) => {
        input.disabled = selected !== "letter";
      });
      if (verification) {
        verification.disabled = selected !== "letter";
        verification.required = selected === "letter";
      }
      if (verificationWrapper) {
        verificationWrapper.hidden = selected !== "letter";
      }
    };
    typeRadios.forEach((radio) => radio.addEventListener("change", update));
    update();
  }
})();
