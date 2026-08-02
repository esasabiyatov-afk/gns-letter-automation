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

  const taxpayerList = document.querySelector("#taxpayer-list");
  const addTaxpayer = document.querySelector("#add-taxpayer");
  if (taxpayerList && addTaxpayer) {
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
    addTaxpayer.addEventListener("click", () => {
      const row = taxpayerList.querySelector(".taxpayer-row").cloneNode(true);
      row.querySelectorAll("input").forEach((input) => { input.value = ""; });
      taxpayerList.appendChild(row);
      bindRemove(row);
    });
    bindRemove(taxpayerList);
  }

  let activeTextInput = null;
  document.querySelectorAll("input[type='text'], input:not([type])").forEach((input) => {
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
    const update = () => {
      const selected = document.querySelector("input[name='page_type']:checked")?.value;
      letterFields.classList.toggle("is-disabled", selected !== "letter");
      letterFields.querySelectorAll("input").forEach((input) => {
        input.disabled = selected !== "letter";
      });
      if (verification) {
        verification.disabled = selected !== "letter";
        verification.required = selected === "letter";
      }
    };
    typeRadios.forEach((radio) => radio.addEventListener("change", update));
    update();
  }
})();
