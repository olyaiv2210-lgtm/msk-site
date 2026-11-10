
(() => {
  "use strict";

  const STORAGE_KEY = "msk_demo_contracts_v1";

  const fields = [
    ["company", "Название организации", true],
    ["inn", "ИНН", true],
    ["kpp", "КПП"],
    ["ogrn", "ОГРН / ОГРНИП"],
    ["okpo", "ОКПО"],
    ["address", "Юридический адрес"],
    ["director", "ФИО руководителя"],
    ["bank", "Банк"],
    ["account", "Расчётный счёт"],
    ["correspondent", "Корреспондентский счёт"],
    ["bik", "БИК"]
  ];

  function loadContracts() {
    try {
      const data = JSON.parse(
        localStorage.getItem(STORAGE_KEY) || "[]"
      );
      return Array.isArray(data) ? data : [];
    } catch {
      return [];
    }
  }

  function saveContracts(data) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  }

  function escapeHTML(value) {
    return String(value ?? "").replace(
      /[&<>"']/g,
      char => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
      })[char]
    );
  }

  function detectType(data) {
    const name = String(data.company || "").trim();
    const inn = String(data.inn || "").replace(/\D/g, "");
    const ogrn = String(data.ogrn || "").replace(/\D/g, "");

    const ip = /^(ИП|Индивидуальный предприниматель)\b/i
      .test(name) || inn.length === 12 || ogrn.length === 15;

    const organization =
      /^(ООО|АО|ПАО)\b/i.test(name) ||
      inn.length === 10 || ogrn.length === 13;

    if (ip && organization) return "unknown";
    if (ip) return "ip";
    if (organization) return "organization";
    return "unknown";
  }

  function extractRequisites(text) {
    const source = String(text || "").replace(/\u00a0/g, " ");
    const lines = source.split(/\r?\n/)
      .map(line => line.trim().replace(/^\d{1,2}\s*[.)]\s*/, ""))
      .filter(Boolean);
    const result = {};

    // Номера ищем по подписям, а не по расположению в документе.
    const numbers = {
      inn: /ИНН\s*[:№-]?\s*(\d{10}|\d{12})(?!\d)/i,
      kpp: /КПП\s*[:№-]?\s*(\d{9})(?!\d)/i,
      ogrn: /ОГРН(?:ИП)?\s*[:№-]?\s*(\d{13}|\d{15})(?!\d)/i,
      okpo: /ОКПО\s*[:№-]?\s*(\d{8}|\d{10})(?!\d)/i,
      bik: /(?:Банковский\s+(?:идентификационный\s+)?код\s*\(?БИК\)?|БИК)\s*[:№-]?\s*(\d{9})(?!\d)/i,
      account: /(?:Номер\s+расч[её]тного\s+сч[её]та|Расч[её]тный\s+сч[её]т|Р\s*\/\s*С)\s*[:№-]?\s*(\d{20})(?!\d)/i,
      correspondent: /(?:Корреспондентский\s+сч[её]т|Корр\.?\s*сч[её]т|К\s*\/\s*С)\s*[:№-]?\s*(\d{20})(?!\d)/i
    };

    for (const [name, pattern] of Object.entries(numbers)) {
      const match = source.match(pattern);
      if (match) result[name] = match[1];
    }

    const labels = {
      company: /^(?:Полное\s+наименование|Наименование\s+(?:предприятия|организации)|Организация)\s*[:\-]\s*(.+)$/i,
      address: /^(?:Юридический\s+адрес|Адрес\s+местонахождения)\s*[:\-]\s*(.+)$/i,
      director: /^(?:ФИО\s+(?:руководителя|директора)|Руководитель|Генеральный\s+директор)\s*[:\-]\s*(.+)$/i,
      bank: /^(?:Наименование\s+банка|Банк)\s*[:\-]\s*(.+)$/i
    };

    for (const line of lines) {
      for (const [name, pattern] of Object.entries(labels)) {
        if (result[name]) continue;
        const match = line.match(pattern);
        if (match) result[name] = match[1].trim();
      }
    }

    // В карточках руководителя часто указывают справа от должности.
    if (!result.director) {
      const directorLine = lines.find(line =>
        /^(?:Генеральный\s+)?директор(?:\s|$)/i.test(line)
      );
      if (directorLine) {
        const fullName = directorLine.match(
          /([А-ЯЁ][а-яё-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.|[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+)\s*$/i
        );
        if (fullName) result.director = fullName[1].replace(/\s+/g, " ").trim();
      }
    }

    if (!result.company) {
      const companyLine = lines.find(line =>
        /^(?:ООО|АО|ПАО|ИП)\s+[«"А-ЯЁ]/i.test(line)
      );
      if (companyLine) result.company = companyLine;
    }

    return result;
  }

  function loadWordLibrary() {
    if (window.mammoth) return Promise.resolve(window.mammoth);

    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src =
        "https://cdn.jsdelivr.net/npm/mammoth@1.9.1/mammoth.browser.min.js";

      script.onload = () => window.mammoth
        ? resolve(window.mammoth)
        : reject(new Error("Библиотека Word недоступна"));

      script.onerror = () =>
        reject(new Error("Нет доступа к библиотеке Word"));

      document.head.appendChild(script);
    });
  }

  async function readCard(file) {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      throw new Error(
        "Пока автоматически распознаются только файлы Word .docx. " +
        "PDF, Excel и фотографии нужно будет подключить отдельно."
      );
    }

    const mammoth = await loadWordLibrary();
    const result = await mammoth.extractRawText({
      arrayBuffer: await file.arrayBuffer()
    });

    if (!result.value.trim()) {
      throw new Error("В документе не найден текст.");
    }

    return extractRequisites(result.value);
  }

  function closeModal() {
    document.getElementById("msk-modal")?.remove();
  }

  function makeModal(content) {
    closeModal();

    const overlay = document.createElement("div");
    overlay.id = "msk-modal";
    overlay.style.cssText = `
      position:fixed;inset:0;z-index:9999;
      background:rgba(8,26,47,.75);
      display:flex;align-items:center;justify-content:center;
      padding:16px;box-sizing:border-box;
    `;

    overlay.innerHTML = `
      <div style="
        background:white;color:#10243e;
        border-radius:18px;padding:24px;
        width:100%;max-width:800px;
        max-height:92vh;overflow:auto;
        box-sizing:border-box;font-family:Arial,sans-serif;
      ">
        ${content}
      </div>
    `;

    overlay.addEventListener("click", event => {
      if (event.target === overlay) closeModal();
    });

    document.body.appendChild(overlay);
    return overlay;
  }

  function openForm(existing = null) {
    const formFields = fields.map(([name, label, required]) => `
      <label style="display:flex;flex-direction:column;gap:6px">
        <span>${label}${required ? " *" : ""}</span>
        <input name="${name}" ${required ? "required" : ""}
          value="${escapeHTML(existing?.[name] || "")}"
          style="padding:11px;border:1px solid #ccd5e0;
          border-radius:8px;box-sizing:border-box;width:100%">
      </label>
    `).join("");

    const overlay = makeModal(`
      <h2>${existing ? "Редактирование" : "Создание"} черновика</h2>
      <p>Тестовый режим. Официальный номер не присваивается.</p>

      <form id="msk-form">
        <div style="padding:16px;background:#eef5fa;
          border-radius:12px;margin-bottom:18px">
          <strong>Карточка контрагента</strong>
          <p>Автоматическое чтение Word .docx</p>
          <input id="msk-upload" type="file"
            accept=".docx">
          <p id="msk-file-status" role="status"></p>
        </div>

        <div style="display:grid;
          grid-template-columns:repeat(auto-fit,minmax(220px,1fr));
          gap:14px">
          ${formFields}
        </div>

        <p id="msk-type" style="margin-top:18px"></p>

        <div style="display:flex;gap:10px;
          justify-content:flex-end;margin-top:20px">
          <button type="button" id="msk-cancel">
            Отмена
          </button>
          <button type="submit" style="
            background:#139c91;color:white;border:0;
            border-radius:8px;padding:12px 18px">
            Сохранить черновик
          </button>
        </div>
      </form>
    `);

    const form = overlay.querySelector("#msk-form");
    const typeLabel = overlay.querySelector("#msk-type");

    function refreshType() {
      const data = Object.fromEntries(new FormData(form));
      const type = detectType(data);

      typeLabel.textContent = type === "ip"
        ? "Тип: индивидуальный предприниматель"
        : type === "organization"
        ? "Тип: организация"
        : "Тип контрагента нужно проверить";
    }

    form.addEventListener("input", refreshType);
    refreshType();

    overlay.querySelector("#msk-upload")
      .addEventListener("change", async event => {
        const file = event.target.files[0];
        if (!file) return;

        const status = overlay.querySelector("#msk-file-status");
        status.textContent = "Читаем карточку...";

        try {
          const data = await readCard(file);
          let count = 0;

          for (const [name, value] of Object.entries(data)) {
            const input = form.elements.namedItem(name);
            if (input && value) {
              input.value = value;
              count++;
            }
          }

          refreshType();
          status.textContent =
            `Заполнено полей: ${count}. Проверьте данные вручную.`;
        } catch (error) {
          status.textContent = error.message;
        }
      });

    overlay.querySelector("#msk-cancel").onclick = closeModal;

    form.addEventListener("submit", event => {
      event.preventDefault();

      const data = Object.fromEntries(new FormData(form));
      const type = detectType(data);

      if (type === "unknown") {
        alert("Проверьте тип контрагента и его реквизиты.");
        return;
      }

      const contracts = loadContracts();
      const now = new Date().toISOString();

      const record = {
        ...data,
        id: existing?.id || crypto.randomUUID(),
        createdAt: existing?.createdAt || now,
        updatedAt: now,
        counterpartyType: type,
        status: "Черновик"
      };

      const index = contracts.findIndex(c => c.id === record.id);

      if (index >= 0) {
        contracts[index] = record;
      } else {
        contracts.push(record);
      }

      try {
        saveContracts(contracts);
        closeModal();
        updateCounter();
        alert("Черновик сохранён в этом браузере.");
      } catch (error) {
        alert("Ошибка сохранения: " + error.message);
      }
    });
  }

  function showRegistry() {
    const contracts = loadContracts();

    const rows = contracts.map((c, index) => `
      <tr>
        <td style="padding:10px;border-bottom:1px solid #ddd">
          ${index + 1}
        </td>
        <td style="padding:10px;border-bottom:1px solid #ddd">
          ${escapeHTML(c.company)}
        </td>
        <td style="padding:10px;border-bottom:1px solid #ddd">
          ${escapeHTML(c.inn)}
        </td>
        <td style="padding:10px;border-bottom:1px solid #ddd">
          ${escapeHTML(c.status)}
        </td>
        <td style="padding:10px;border-bottom:1px solid #ddd">
          <button type="button" data-edit="${escapeHTML(c.id)}">
            Изменить
          </button>
        </td>
      </tr>
    `).join("");

    const overlay = makeModal(`
      <h2>Реестр тестовых черновиков</h2>
      <p>Записи хранятся только в этом браузере.</p>

      <div style="overflow-x:auto">
        <table style="width:100%;border-collapse:collapse">
          <thead>
            <tr>
              <th>№</th><th>Контрагент</th>
              <th>ИНН</th><th>Статус</th><th>Действия</th>
            </tr>
          </thead>
          <tbody>
            ${rows || '<tr><td colspan="5">Записей пока нет</td></tr>'}
          </tbody>
        </table>
      </div>

      <button id="msk-close" style="margin-top:20px">
        Закрыть
      </button>
    `);

    overlay.querySelector("#msk-close").onclick = closeModal;

    overlay.addEventListener("click", event => {
      const button = event.target.closest("[data-edit]");
      if (!button) return;

      const record = contracts.find(c => c.id === button.dataset.edit);
      if (record) openForm(record);
    });
  }

  function updateCounter() {
    const count = loadContracts().length;
    const counter = document.getElementById("count");

    if (counter) counter.textContent = String(count);
  }

  document.addEventListener("click", event => {
    const target = event.target.closest("button,a");
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
