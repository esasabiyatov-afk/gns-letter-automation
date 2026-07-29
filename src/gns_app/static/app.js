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

  const reviewImage = document.querySelector("#review-image");
  if (reviewImage) {
    let zoom = 1;
    let enhanced = false;
    document.querySelectorAll("[data-zoom]").forEach((button) => {
      button.addEventListener("click", () => {
        const delta = Number(button.dataset.zoom);
        zoom = delta === 0 ? 1 : Math.min(2.5, Math.max(0.55, zoom + delta));
        reviewImage.style.width = `${zoom * 100}%`;
      });
    });
    document.querySelector("#toggle-image")?.addEventListener("click", (event) => {
      enhanced = !enhanced;
      reviewImage.src = enhanced
        ? reviewImage.dataset.enhanced
        : reviewImage.dataset.original;
      event.currentTarget.textContent = enhanced ? "Оригинал" : "Улучшенная";
    });
  }

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
    const update = () => {
      const selected = document.querySelector("input[name='page_type']:checked")?.value;
      letterFields.classList.toggle("is-disabled", selected !== "letter");
      letterFields.querySelectorAll("input").forEach((input) => {
        input.disabled = selected !== "letter";
      });
    };
    typeRadios.forEach((radio) => radio.addEventListener("change", update));
    update();
  }
})();

