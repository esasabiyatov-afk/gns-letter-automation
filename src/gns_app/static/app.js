(() => {
  const employeeEntry = document.querySelector("#employee-entry");
  if (employeeEntry && window.location.hash === "#employee-entry") {
    employeeEntry.open = true;
  }
  document.querySelector(".active-employee-chip")?.addEventListener("click", () => {
    if (employeeEntry) employeeEntry.open = true;
  });

  const tourCard = document.querySelector("[data-tour-card]");
  const tourStart = document.querySelector("[data-tour-start]");
  if (tourCard && tourStart) {
    const tourSteps = [
      {
        id: "work-stages",
        path: "/?tab=incoming",
        selector: "[data-tour-anchor='work-tabs']",
        title: "Работайте слева направо",
        text: "Основной маршрут состоит из трёх этапов: «Входящие» → «Проверка» → «Ответы». Числа показывают, сколько работы сейчас находится на каждом этапе.",
      },
      {
        id: "receive-mail",
        path: "/?tab=incoming",
        selector: "[data-tour-anchor='incoming']",
        title: "Получите входящие письма",
        text: "Нажмите «Получить письма»: PDF загрузятся из Outlook, а одинаковые файлы не будут обработаны повторно. Ручная загрузка и проверка папки находятся рядом.",
      },
      {
        id: "review",
        path: "/?tab=review",
        selector: "[data-tour-anchor='review']",
        title: "Проверьте только сомнения",
        text: "В этой очереди появляются только решения, которые приложение не имеет права принимать само: проблемные страницы, расхождения, АБС и ОДБ.",
      },
      {
        id: "prepare",
        path: "/?tab=responses&response_view=prepare",
        selector: "[data-tour-anchor='prepare']",
        title: "Создайте готовые ответы",
        text: "Во вкладке «К созданию» сформируйте один Word для срочного письма или сразу весь готовый пакет.",
      },
      {
        id: "created",
        path: "/?tab=responses&response_view=created",
        selector: "[data-tour-anchor='created']",
        title: "Присвойте номера и отсканируйте",
        text: "Введите первый свободный исходящий номер — остальные заполнятся по порядку. Затем откройте Word, соберите подпись и печать и нажмите «Сканировать» у нужного письма.",
      },
      {
        id: "manual",
        path: "/?tab=responses&response_view=manual",
        selector: "[data-tour-anchor='manual']",
        title: "Не пропустите ручные ответы",
        text: "Если подтверждено наличие расчётного счёта, обращение не попадёт в шаблон «счета отсутствуют» и останется здесь для отдельного ответа.",
      },
      {
        id: "history",
        path: "/history",
        selector: "[data-tour-anchor='history']",
        title: "Найдите любое письмо",
        text: "История ищет по файлу, лицу, ИНН, отправителю и исходящему номеру. Связанные PDF, Word и сканы открываются через меню «⋯».",
      },
    ];
    const storagePrefix = "gns-guided-tour-v1";
    const activeKey = `${storagePrefix}:active`;
    const stepKey = `${storagePrefix}:step`;
    const seenKey = `${storagePrefix}:seen`;
    const progress = tourCard.querySelector("[data-tour-progress]");
    const title = tourCard.querySelector("[data-tour-title]");
    const text = tourCard.querySelector("[data-tour-text]");
    const back = tourCard.querySelector("[data-tour-back]");
    const next = tourCard.querySelector("[data-tour-next]");
    const skip = tourCard.querySelector("[data-tour-skip]");
    const close = tourCard.querySelector("[data-tour-close]");
    let activeTarget = null;
    let volatileActive = false;
    let volatileStep = 0;
    let volatileSeen = false;

    const sessionGet = (key) => {
      try {
        return window.sessionStorage.getItem(key);
      } catch (_error) {
        if (key === activeKey) return volatileActive ? "1" : null;
        if (key === stepKey) return String(volatileStep);
        return null;
      }
    };
    const sessionSet = (key, value) => {
      if (key === activeKey) volatileActive = value === "1";
      if (key === stepKey) volatileStep = Number.parseInt(value, 10) || 0;
      try {
        window.sessionStorage.setItem(key, value);
      } catch (_error) {
        // В пределах страницы используется резервное состояние в памяти.
      }
    };
    const sessionRemove = (key) => {
      if (key === activeKey) volatileActive = false;
      if (key === stepKey) volatileStep = 0;
      try {
        window.sessionStorage.removeItem(key);
      } catch (_error) {
        // Нечего очищать.
      }
    };
    const completedGet = () => {
      try {
        return window.localStorage.getItem(seenKey);
      } catch (_error) {
        return volatileSeen ? "1" : null;
      }
    };
    const completedSet = () => {
      volatileSeen = true;
      try {
        window.localStorage.setItem(seenKey, "1");
      } catch (_error) {
        // Повторный автоматический показ подавляется в пределах страницы.
      }
    };
    const clearTarget = () => {
      activeTarget?.classList.remove("tour-target-active");
      activeTarget = null;
    };
    const routeMatches = (path) => {
      const expected = new URL(path, window.location.origin);
      const current = new URL(window.location.href);
      if (expected.pathname !== current.pathname) return false;
      const expectedTab = expected.searchParams.get("tab");
      const currentTab = current.searchParams.get("tab")
        || (current.pathname === "/" ? "incoming" : "");
      const expectedView = expected.searchParams.get("response_view");
      if (expectedTab && expectedTab !== currentTab) return false;
      if (expectedView
          && expectedView !== current.searchParams.get("response_view")) {
        return false;
      }
      return true;
    };
    const finishTour = () => {
      clearTarget();
      tourCard.hidden = true;
      document.body.classList.remove("tour-is-active");
      completedSet();
      sessionRemove(activeKey);
      sessionRemove(stepKey);
    };
    const renderTour = (allowNavigation = false) => {
      const storedStep = Number.parseInt(sessionGet(stepKey) || "0", 10);
      const index = Number.isFinite(storedStep)
        ? Math.min(Math.max(storedStep, 0), tourSteps.length - 1)
        : 0;
      const step = tourSteps[index];
      if (!routeMatches(step.path)) {
        clearTarget();
        tourCard.hidden = true;
        document.body.classList.remove("tour-is-active");
        if (allowNavigation) window.location.assign(step.path);
        return;
      }

      clearTarget();
      activeTarget = document.querySelector(step.selector);
      activeTarget?.classList.add("tour-target-active");
      if (activeTarget) {
        const rectangle = activeTarget.getBoundingClientRect();
        if (rectangle.top < 70 || rectangle.bottom > window.innerHeight - 150) {
          activeTarget.scrollIntoView({ behavior: "smooth", block: "center" });
        }
      }

      progress.textContent = `${index + 1} из ${tourSteps.length}`;
      title.textContent = step.title;
      text.textContent = step.text;
      back.hidden = index === 0;
      next.textContent = index === tourSteps.length - 1 ? "Завершить" : "Далее";
      tourCard.hidden = false;
      document.body.classList.add("tour-is-active");
    };
    const goToStep = (index) => {
      const bounded = Math.min(Math.max(index, 0), tourSteps.length - 1);
      sessionSet(activeKey, "1");
      sessionSet(stepKey, String(bounded));
      renderTour(true);
    };
    const currentStep = () => Number.parseInt(sessionGet(stepKey) || "0", 10);

    tourStart.addEventListener("click", () => {
      goToStep(sessionGet(activeKey) === "1" ? currentStep() : 0);
    });
    back?.addEventListener("click", () => goToStep(currentStep() - 1));
    next?.addEventListener("click", () => {
      const index = currentStep();
      if (index >= tourSteps.length - 1) {
        finishTour();
      } else {
        goToStep(index + 1);
      }
    });
    skip?.addEventListener("click", finishTour);
    close?.addEventListener("click", finishTour);
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && !tourCard.hidden) finishTour();
    });

    const firstWorkScreen = window.location.pathname === "/"
      && [null, "incoming"].includes(
        new URL(window.location.href).searchParams.get("tab")
      );
    if (sessionGet(activeKey) === "1") {
      renderTour(false);
    } else if (completedGet() !== "1" && firstWorkScreen) {
      window.setTimeout(() => goToStep(0), 350);
    }
  }

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

  document.querySelectorAll("[data-auto-submit]").forEach((select) => {
    select.addEventListener("change", () => {
      select.closest("form[data-auto-submit-form]")?.requestSubmit();
    });
  });

  document.querySelectorAll("[data-open-dialog]").forEach((button) => {
    button.addEventListener("click", () => {
      document.getElementById(button.dataset.openDialog)?.showModal();
    });
  });
  document.querySelectorAll("[data-close-dialog]").forEach((button) => {
    button.addEventListener("click", () => button.closest("dialog")?.close());
  });

  const inlineNumberForms = [
    ...document.querySelectorAll("form[data-inline-number-form]"),
  ];
  if (inlineNumberForms.length) {
    const entries = inlineNumberForms.map((form) => ({
      form,
      input: form.querySelector("[data-inline-number-input]"),
      submit: form.querySelector("[data-inline-number-submit]"),
      letterId: form.dataset.letterId || "",
      originalAction: form.getAttribute("action") || "",
    })).filter((entry) => entry.input && entry.submit && entry.letterId);
    const unnumbered = entries.filter(
      (entry) => (entry.input.dataset.originalNumber || "") === "",
    );

    const addSequenceField = (form, name, value) => {
      const field = document.createElement("input");
      field.type = "hidden";
      field.name = name;
      field.value = value;
      field.dataset.inlineSequenceField = "";
      form.append(field);
    };

    const resetSequence = (source, typedValue) => {
      entries.forEach((entry) => {
        entry.form.querySelectorAll("[data-inline-sequence-field]")
          .forEach((field) => field.remove());
        entry.form.setAttribute("action", entry.originalAction);
        entry.form.classList.remove("sequence-source");
        entry.submit.disabled = false;
        entry.submit.textContent = "✓";
        entry.submit.title = "Сохранить номер";
        if (entry.input.dataset.sequencePreview === "true") {
          if (entry !== source) entry.input.value = "";
          entry.input.classList.remove("sequence-preview");
          delete entry.input.dataset.sequencePreview;
        }
      });
      if (source) source.input.value = typedValue;
    };

    const renderSequence = (source) => {
      const typedValue = source.input.value.trim();
      resetSequence(source, typedValue);
      const startIndex = unnumbered.indexOf(source);
      if (startIndex < 0 || !/^\d{1,9}$/.test(typedValue)) return;
      const firstNumber = Number.parseInt(typedValue, 10);
      const suffix = unnumbered.slice(startIndex);
      if (
        firstNumber < 1
        || firstNumber + Math.max(0, suffix.length - 1) > 999999999
      ) return;

      suffix.forEach((entry, offset) => {
        entry.input.value = String(firstNumber + offset);
        entry.input.classList.add("sequence-preview");
        entry.input.dataset.sequencePreview = "true";
        if (entry !== source) {
          entry.submit.disabled = true;
          entry.submit.title = "Сохранится вместе с первым номером";
        }
      });
      source.form.setAttribute("action", "/today/outgoing-numbers/assign");
      source.form.classList.add("sequence-source");
      addSequenceField(source.form, "first_number", String(firstNumber));
      suffix.forEach((entry) => {
        addSequenceField(source.form, "letter_ids", entry.letterId);
      });
      source.submit.textContent = "✓";
      source.submit.title = suffix.length > 1
        ? `Сохранить номера для ${suffix.length} писем`
        : "Сохранить номер";
    };

    unnumbered.forEach((entry) => {
      entry.input.addEventListener("input", () => renderSequence(entry));
    });
  }

  const actionMenus = [...document.querySelectorAll(".action-menu")];
  actionMenus.forEach((menu) => {
    menu.addEventListener("toggle", () => {
      if (!menu.open) return;
      actionMenus.forEach((other) => {
        if (other !== menu) other.open = false;
      });
    });
  });
  document.addEventListener("click", (event) => {
    actionMenus.forEach((menu) => {
      if (menu.open && !menu.contains(event.target)) menu.open = false;
    });
  });

  const workflowTabs = [...document.querySelectorAll("[data-workflow-tab]")];
  const workflowPanels = [...document.querySelectorAll("[data-workflow-panel]")];
  if (workflowTabs.length && workflowPanels.length) {
    const activateWorkflowTab = (name) => {
      workflowTabs.forEach((button) => {
        const active = button.dataset.workflowTab === name;
        button.classList.toggle("active", active);
        button.setAttribute("aria-selected", active ? "true" : "false");
      });
      workflowPanels.forEach((panel) => {
        panel.hidden = panel.dataset.workflowPanel !== name;
      });
    };
    const hashTarget = window.location.hash
      ? document.getElementById(window.location.hash.slice(1))
      : null;
    const hashPanel = hashTarget?.closest("[data-workflow-panel]");
    const query = new URLSearchParams(window.location.search);
    const initialTab = hashPanel?.dataset.workflowPanel
      || (query.has("outgoing_start") ? "created" : "prepare");
    activateWorkflowTab(initialTab);
    workflowTabs.forEach((button) => {
      button.addEventListener("click", () => {
        activateWorkflowTab(button.dataset.workflowTab);
      });
    });
  }

  document.querySelectorAll(".response-group-card details").forEach((details) => {
    details.addEventListener("toggle", () => {
      const card = details.closest(".response-group-card");
      if (!card) return;
      card.classList.toggle(
        "is-expanded",
        Boolean(card.querySelector("details[open]"))
      );
    });
  });

  const settingsTabs = [...document.querySelectorAll("[data-settings-tab]")];
  const settingsPanels = [...document.querySelectorAll("[data-settings-panel]")];
  const settingsSaveState = document.querySelector("[data-settings-save-state]");
  if (settingsTabs.length && settingsPanels.length) {
    const sectionNames = settingsTabs.map((button) => button.dataset.settingsTab);
    const selectSettingsSection = (name, updateHash = false) => {
      const selected = sectionNames.includes(name) ? name : sectionNames[0];
      settingsTabs.forEach((button) => {
        const active = button.dataset.settingsTab === selected;
        button.setAttribute("aria-selected", active ? "true" : "false");
      });
      settingsPanels.forEach((panel) => {
        panel.hidden = panel.dataset.settingsPanel !== selected;
      });
      if (updateHash) {
        window.history.replaceState(null, "", `#${selected}`);
      }
    };
    selectSettingsSection(window.location.hash.slice(1));
    settingsTabs.forEach((button) => {
      button.addEventListener("click", () => {
        selectSettingsSection(button.dataset.settingsTab, true);
      });
    });
    window.addEventListener("hashchange", () => {
      selectSettingsSection(window.location.hash.slice(1));
    });
  }

  let settingsDirty = false;
  document.querySelectorAll("[data-settings-form]").forEach((form) => {
    const markDirty = () => {
      settingsDirty = true;
      if (settingsSaveState) {
        settingsSaveState.textContent = "Есть несохранённые изменения";
        settingsSaveState.classList.add("is-dirty");
        settingsSaveState.classList.remove("is-saving");
      }
    };
    form.addEventListener("input", markDirty);
    form.addEventListener("change", markDirty);
    form.addEventListener("submit", () => {
      settingsDirty = false;
      if (settingsSaveState) {
        settingsSaveState.textContent = "Сохраняем…";
        settingsSaveState.classList.remove("is-dirty");
        settingsSaveState.classList.add("is-saving");
      }
    });
  });
  window.addEventListener("beforeunload", (event) => {
    if (!settingsDirty) return;
    event.preventDefault();
    event.returnValue = "";
  });

  document.querySelectorAll("[data-paged-list]").forEach((list) => {
    const items = [...list.children].filter((item) => item.hasAttribute("data-page-item"));
    const configuredPageSize = Math.max(
      1,
      Number.parseInt(list.dataset.pageSize || "6", 10)
    );
    const responsivePageSize = window.matchMedia("(max-width: 720px)").matches
      ? 1
      : window.matchMedia("(max-width: 1050px)").matches
        ? 2
        : configuredPageSize;
    const pageSize = Math.min(configuredPageSize, responsivePageSize);
    const pageCount = Math.ceil(items.length / pageSize);
    if (pageCount <= 1) return;
    let page = 0;
    const controls = document.createElement("nav");
    controls.className = "pagination-controls";
    controls.setAttribute("aria-label", "Страницы списка");
    const previous = document.createElement("button");
    previous.type = "button";
    previous.textContent = "←";
    previous.setAttribute("aria-label", "Предыдущая страница списка");
    const indicator = document.createElement("strong");
    const next = document.createElement("button");
    next.type = "button";
    next.textContent = "→";
    next.setAttribute("aria-label", "Следующая страница списка");
    controls.append(previous, indicator, next);
    list.after(controls);
    const render = () => {
      items.forEach((item, index) => {
        item.hidden = index < page * pageSize || index >= (page + 1) * pageSize;
      });
      indicator.textContent = `${page + 1} из ${pageCount}`;
      previous.disabled = page === 0;
      next.disabled = page === pageCount - 1;
    };
    previous.addEventListener("click", () => {
      if (page > 0) page -= 1;
      render();
    });
    next.addEventListener("click", () => {
      if (page + 1 < pageCount) page += 1;
      render();
    });
    render();
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
          const willBeEmpty = !odbCase.querySelector("[data-odb-row]");
          const panel = odbCase.closest(".panel");
          if (willBeEmpty) odbCase.remove();
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
    const selectViewerTab = (button) => {
      const selected = button.dataset.viewerTab;
      viewerTabs.forEach((tab) => {
        const active = tab === button;
        tab.classList.toggle("active", active);
        tab.setAttribute("aria-pressed", active ? "true" : "false");
      });
      viewerPanels.forEach((panel) => {
        panel.hidden = panel.dataset.viewerPanel !== selected;
      });
    };
    viewerTabs.forEach((button) => {
      button.addEventListener("click", () => selectViewerTab(button));
    });
    document.addEventListener("keydown", (event) => {
      if (!event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return;
      const index = { "1": 0, "2": 1, "3": 2 }[event.key];
      if (index === undefined || !viewerTabs[index]) return;
      event.preventDefault();
      selectViewerTab(viewerTabs[index]);
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
    taxpayerList.querySelectorAll("[data-ocr-inn-candidate]").forEach((choice) => {
      choice.addEventListener("change", () => {
        const innInput = taxpayerList.querySelector("input[name='taxpayer_inn']");
        if (innInput && choice.checked) {
          innInput.value = choice.value || "";
          innInput.dispatchEvent(new Event("input", { bubbles: true }));
          innInput.focus();
        }
      });
    });
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
        message.textContent = `${data.provider || "Реестр"}: ${data.official_name}`;
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
        result.textContent = "Проверяем ИНН в реестре…";
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

  const selectedUploadPage = document.querySelector(
    ".upload-page-row.is-selected"
  );
  if (selectedUploadPage) {
    window.requestAnimationFrame(() => {
      selectedUploadPage.scrollIntoView({ block: "nearest", inline: "nearest" });
    });
  }
})();
