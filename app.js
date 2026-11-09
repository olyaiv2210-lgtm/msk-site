

(() => {
  "use strict";

  const STORAGE_KEY = "msk_demo_contracts_v1";
  
  async function loadWordLibrary() {
    if (window.mammoth) return window.mammoth;

    const script = document.createElement("script");
    script.src =
      "https://cdn.jsdelivr.net/npm/mammoth@1.9.1/mammoth.browser.min.js";

    return new Promise((resolve, reject) => {
      script.onload = () => {
        if (window.mammoth) {
          resolve(window.mammoth);
        } else {
          reject(new Error("Библиотека Word не загрузилась"));
        }
      };

      script.onerror = () => {
        reject(new Error("Не удалось загрузить библиотеку Word"));
      };

      document.head.appendChild(script);
    });
  }


  
  function extractRequisites(text) {
    const result = {};

    const patterns = {
      inn: /\bИНН\s*[:№-]?\s*(\d{10}|\d{12})\b/i,
      kpp: /\bКПП\s*[:№-]?\s*(\d{9})\b/i,
      ogrn: /\bОГРН(?:ИП)?\s*[:№-]?\s*(\d{13}|\d{15})\b/i,
      okpo: /\bОКПО\s*[:№-]?\s*(\d{8}|\d{10})\b/i,
      bik: /\bБИК\s*[:№-]?\s*(\d{9})\b/i,
      correspondent:
        /(?:Корр(?:еспондентский)?\.?\s*сч[её]т|К\/С)\s*[:№-]?\s*(\d{20})\b/i,
      account:
        /(?:Расч[её]тный\s*сч[её]т|Р\/С)\s*[:№-]?\s*(\d{20})\b/i
    };

    for (const [field, pattern] of Object.entries(patterns)) {
      const match = text.match(pattern);
      if (match) result[field] = match[1];
    }

    return result;
  }

  function fillRequisites(form, requisites) {
    let count = 0;

    for (const [name, value] of Object.entries(requisites)) {
      const input = form.elements.namedItem(name);

      if (input && value) {
        input.value = value;
        count++;
      }
    }

    return count;
  }


  function loadContracts() {
    try {
      const data = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
      return Array.isArray(data) ? data : [];
    } catch {
      return [];
    }
  }

  function saveContracts(contracts) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(contracts));
  }

  function escapeHTML(value) {
    return String(value ?? "").replace(/[&<>"']/g, char => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    })[char]);
  }

  const fields = [
    ["company", "Название организации", true],
    ["inn", "ИНН", true],
    ["kpp", "КПП"],
    ["ogrn", "ОГРН"],
    ["okpo", "ОКПО"],
    ["address", "Юридический адрес"],
    ["director", "ФИО руководителя"],
    ["bank", "Наименование банка"],
    ["account", "Расчётный счёт"],
    ["correspondent", "Корреспондентский счёт"],
    ["bik", "БИК"]
  ];

  function openForm() {
    document.getElementById("msk-modal")?.remove();

    const overlay = document.createElement("div");
    overlay.id = "msk-modal";
    overlay.style.cssText = `
      position:fixed;inset:0;z-index:9999;
      background:rgba(10,28,50,.75);
      display:flex;align-items:center;justify-content:center;
      padding:15px;box-sizing:border-box;
    `;

    overlay.innerHTML = `
      <div style="background:white;border-radius:18px;
        width:100%;max-width:750px;max-height:92vh;
        overflow-y:auto;padding:28px;box-sizing:border-box;
        font-family:Arial,sans-serif;color:#10243e">

        <h2 style="margin-top:0">Создание договора ЭМСК</h2>
        <p style="color:#65758a">
          Заполните реквизиты контрагента.
          Пока это тестовый черновик без номера договора.
        </p>

        <form id="msk-contract-form">
        
<div style="background:#f0f6fc;padding:18px;
  border:2px dashed #a8c5e0;border-radius:12px;
  margin-bottom:22px">

  <div style="font-weight:bold;margin-bottom:8px">
    📎 Карточка контрагента
  </div>

  <p style="font-size:13px;color:#65758a">
    Выберите файл PDF, Word, Excel или фотографию.
  </p>

  <input type="file" id="msk-card-file"
    accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg,.png,.webp"
    style="width:100%;max-width:100%">

  <p style="font-size:12px;color:#65758a">
    Пока файл только выбирается.
    Автоматическое распознавание подключим позже.
  </p>

</div>

          <div style="display:grid;
            grid-template-columns:repeat(auto-fit,minmax(220px,1fr));
            gap:15px">

            ${fields.map(([name, label, required]) => `
              <label style="display:flex;flex-direction:column;gap:6px">
                <span style="font-size:14px;font-weight:bold">
                  ${label}${required ? " *" : ""}
                </span>
                <input name="${name}" ${required ? "required" : ""}
                  style="padding:12px;border:1px solid #ccd5e0;
                  border-radius:9px;font-size:15px;width:100%;
                  box-sizing:border-box">
              </label>
            `).join("")}

          </div>

          <div style="display:flex;justify-content:flex-end;
            gap:10px;margin-top:25px;flex-wrap:wrap">

            <button type="button" id="msk-cancel"
              style="padding:12px 20px;border:1px solid #ccd5e0;
              border-radius:9px;background:white;cursor:pointer">
              Отмена
            </button>

            <button type="submit"
              style="padding:12px 20px;border:0;border-radius:9px;
              background:#1768a6;color:white;font-weight:bold;
              cursor:pointer">
              Сохранить черновик
            </button>
          </div>
        </form>
      </div>
    `;

    document.body.appendChild(overlay);
    
    overlay.querySelector("#msk-card-file")
      .addEventListener("change", async event => {
        const file = event.target.files[0];
        if (!file) return;

        if (!file.name.toLowerCase().endsWith(".docx")) {
          alert("Пока автоматически читаем только Word .docx");
          return;
        }

        try {
          const mammoth = await loadWordLibrary();
          const buffer = await file.arrayBuffer();
          const result = await mammoth.extractRawText({
            arrayBuffer: buffer
          });

          const text = result.value.trim();

          if (!text) {
            alert("Не удалось найти текст в документе.");
            return;
          }

             
          const form = overlay.querySelector("#msk-contract-form");
          const requisites = extractRequisites(text);
          const count = fillRequisites(form, requisites);

          alert(
            "Документ Word успешно прочитан!\n\n" +
            "Автоматически заполнено полей: " + count +
            "\n\nПроверьте все реквизиты перед сохранением."
          );

        } catch (error) {
          alert("Ошибка чтения Word: " + error.message);
        }
      });


    overlay.querySelector("#msk-cancel").onclick = () => overlay.remove();

    overlay.querySelector("#msk-contract-form").onsubmit = event => {
      event.preventDefault();

      const data = Object.fromEntries(
        new FormData(event.currentTarget).entries()
      );

      const contracts = loadContracts();

      contracts.push({
        id: crypto.randomUUID(),
        ...data,
        createdAt: new Date().toISOString(),
        status: "Черновик"
      });

      try {
        saveContracts(contracts);
        overlay.remove();
        updateCounter();
        alert("Черновик успешно сохранён в этом браузере!");
      } catch (error) {
        alert("Ошибка сохранения: " + error.message);
      }
    };
  }

  function showRegistry() {
    const contracts = loadContracts();

    document.getElementById("msk-modal")?.remove();

    const overlay = document.createElement("div");
    overlay.id = "msk-modal";
    overlay.style.cssText = `
      position:fixed;inset:0;z-index:9999;
      background:rgba(10,28,50,.75);
      display:flex;align-items:center;justify-content:center;
      padding:15px;box-sizing:border-box;
    `;

    const rows = contracts.map((c, i) => `
      <tr>
        <td style="padding:10px;border-bottom:1px solid #ddd">${i + 1}</td>
        <td style="padding:10px;border-bottom:1px solid #ddd">
          ${escapeHTML(c.company)}
        </td>
        <td style="padding:10px;border-bottom:1px solid #ddd">
          ${escapeHTML(c.inn)}
        </td>
        <td style="padding:10px;border-bottom:1px solid #ddd">
          ${escapeHTML(c.status)}
        </td>
      </tr>
    `).join("");

    overlay.innerHTML = `
      <div style="background:white;border-radius:18px;
        width:100%;max-width:850px;max-height:90vh;
        overflow:auto;padding:25px;box-sizing:border-box;
        font-family:Arial,sans-serif;color:#10243e">

        <h2>Реестр тестовых черновиков ЭМСК</h2>

        ${contracts.length ? `
          <div style="overflow-x:auto">
            <table style="width:100%;border-collapse:collapse;text-align:left">
              <thead>
                <tr>
                  <th>№</th>
                  <th>Контрагент</th>
                  <th>ИНН</th>
                  <th>Статус</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        ` : "<p>Черновиков пока нет.</p>"}

        <p style="color:#65758a;font-size:13px">
          Это локальные тестовые записи, а не официальный реестр договоров.
        </p>

        <button id="msk-close"
          style="padding:12px 20px;border:0;border-radius:9px;
          background:#1768a6;color:white;cursor:pointer">
          Закрыть
        </button>
      </div>
    `;

    document.body.appendChild(overlay);
    overlay.querySelector("#msk-close").onclick = () => overlay.remove();
  }

  function updateCounter() {
    const count = loadContracts().length;

    document.querySelectorAll("*").forEach(el => {
      if (el.children.length || el.textContent.trim() !==
          "Черновики в браузере") return;

      const card = el.parentElement;
      const value = [...card.querySelectorAll("*")].find(node =>
        !node.children.length && /^\d+$/.test(node.textContent.trim())
      );

      if (value) value.textContent = String(count);
    });
  }

  document.addEventListener("click", event => {
    const target = event.target.closest("button, a");
    if (!target || target.closest("#msk-modal")) return;

    const label = target.textContent.trim().toLowerCase();

    if (label.includes("создать договор")) {
      event.preventDefault();
      openForm();
    } else if (label.includes("реестр договоров")) {
      event.preventDefault();
      showRegistry();
    }
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", updateCounter);
  } else {
    updateCounter();
  }
})();
