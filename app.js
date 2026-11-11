
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
    ["bik", "БИК"],
    ["actualAddress", "Фактический адрес"],
    ["okved", "ОКВЭД"],
    ["phone", "Телефон"],
    ["phone2", "Дополнительный телефон"],
    ["email", "E-mail"],
    ["edo", "ID ЭДО"]
  ];

  function loadContracts() {
    try {
      const data = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
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

    const ip =
      /^(ИП|Индивидуальный предприниматель)(?=\s|$)/i.test(name) ||
      inn.length === 12 ||
      ogrn.length === 15;

    const organization =
      /^(ООО|АО|ПАО)(?=\s|$)/i.test(name) ||
      inn.length === 10 ||
      ogrn.length === 13;

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

    if (!result.director) {
      const directorLine = lines.find(line =>
        /^(?:Генеральный\s+)?директор(?:\s|$)/i.test(line)
      );

      if (directorLine) {
        const fullName = directorLine.match(
          /([А-ЯЁ][а-яё-]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.|[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+)\s*$/i
        );
        if (fullName) {
          result.director = fullName[1].replace(/\s+/g, " ").trim();
        }
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

  function loadScript(url, globalName) {
    if (window[globalName]) {
      return Promise.resolve(window[globalName]);
    }

    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = url;

      script.onload = () => window[globalName]
        ? resolve(window[globalName])
        : reject(new Error("Библиотека " + globalName + " не загрузилась"));

      script.onerror = () =>
        reject(new Error("Не удалось загрузить " + globalName));

      document.head.appendChild(script);
    });
  }

  async function readCard(file) {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      throw new Error(
        "Пока автоматически распознаются только файлы Word .docx."
      );
    }

    const mammoth = await loadScript(
      "https://cdn.jsdelivr.net/npm/mammoth@1.9.1/mammoth.browser.min.js",
      "mammoth"
    );

    const result = await mammoth.extractRawText({
      arrayBuffer: await file.arrayBuffer()
    });

    if (!result.value.trim()) {
      throw new Error("В документе не найден текст.");
    }

    return extractRequisites(result.value);
  }

  function shortName(name) {
    return String(name || "")
      .replace(/^(?:ИП|Индивидуальный предприниматель)\s+/i, "")
      .trim();
  }

  function initials(name) {
    const parts = shortName(name).split(/\s+/);
    if (parts.length < 3) return shortName(name);
    return parts[0] + " " + parts[1][0] + "." + parts[2][0] + ".";
  }

  async function downloadWord(record, statusElement) {
    const type = detectType(record);

    if (type === "unknown") {
      throw new Error("Не удалось определить тип контрагента.");
    }

    const template = type === "ip"
      ? "template-ip.docx"
      : "template-ooo.docx";

    statusElement.textContent = "Загружаем шаблон " + template + "...";

    const response = await fetch(template, { cache: "no-store" });

    if (!response.ok) {
      throw new Error(
        "Не найден " + template +
        ". Проверьте наличие файла рядом с приложением."
      );
    }

    const buffer = await response.arrayBuffer();

    const PizZip = await loadScript(
      "https://cdn.jsdelivr.net/npm/pizzip@3.2.0/dist/pizzip.min.js",
      "PizZip"
    );

    const Docxtemplater = await loadScript(
      "https://cdn.jsdelivr.net/npm/docxtemplater@3.54.0/build/docxtemplater.js",
      "docxtemplater"
    );

    const zip = new PizZip(buffer);

    const documentXml =
      zip.file("word/document.xml")?.asText() || "";

    if (
      !documentXml.includes("{{") &&
      !documentXml.includes("<w:t>{")
    ) {
      throw new Error(
        "В шаблоне нет полей {{...}}. " +
        "Сначала нужно подготовить шаблон, " +
        "не меняя утверждённый текст договора."
      );
    }

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "{{", end: "}}" },
      nullGetter: () => ""
    });

    const date = new Date().toLocaleDateString("ru-RU", {
      day: "2-digit",
      month: "long",
      year: "numeric"
    });

    const data = {
      НОМЕР_ДОГОВОРА: "ПРОЕКТ (без номера)",
      ДАТА_ДОГОВОРА: date,
      ИП_ФИО: shortName(record.company),
      ИП_ИНИЦИАЛЫ: initials(record.company),
      ОГРНИП: record.ogrn || "",
      ИНН: record.inn || "",
      КПП: record.kpp || "",
      ОГРН: record.ogrn || "",
      ОКПО: record.okpo || "",
      АДРЕС_РЕГИСТРАЦИИ: record.address || "",
      АДРЕС_ФАКТИЧЕСКИЙ: record.actualAddress || "",
      БАНК: record.bank || "",
      РАСЧЕТНЫЙ_СЧЕТ: record.account || "",
      КОРР_СЧЕТ: record.correspondent || "",
      БИК: record.bik || "",
      ОКВЭД: record.okved || "",
      ТЕЛЕФОН_1: record.phone || "",
      ТЕЛЕФОН_2: record.phone2 || "",
      EMAIL: record.email || "",
      ID_ЭДО: record.edo || "",
      НАЗВАНИЕ_ОРГАНИЗАЦИИ: record.company || "",
      ФИО_РУКОВОДИТЕЛЯ: record.director || ""
    };

    try {
      doc.render(data);
    } catch (error) {
      throw new Error(
        "Ошибка полей шаблона Word: " + (error.message || error)
      );
    }

    const blob = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");

    a.href = url;
    a.download =
      "ПРОЕКТ_" + (record.inn || "контрагент") + ".docx";

    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => URL.revokeObjectURL(url), 30000);

    statusElement.textContent =
      "Word-проект скачан. Официальный номер не присвоен.";
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
          <input id="msk-upload" type="file" accept=".docx">
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
          <button type="button" data-word="${escapeHTML(c.id)}">
            Скачать Word
          </button>
        </td>
      </tr>
    `).join("");

    const overlay = makeModal(`
      <h2>Реестр тестовых черновиков</h2>
      <p>
        Записи хранятся только в этом браузере.
        Word — тестовый проект без официального номера.
      </p>

      <p id="msk-word-status" role="status"></p>

      <div style="overflow-x:auto">
        <table style="width:100%;border-collapse:collapse">
          <thead>
            <tr>
              <th>№</th>
              <th>Контрагент</th>
              <th>ИНН</th>
              <th>Статус</th>
              <th>Действия</th>
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

    overlay.addEventListener("click", async event => {
      const editButton = event.target.closest("[data-edit]");

      if (editButton) {
        const record = contracts.find(
          c => c.id === editButton.dataset.edit
        );

        if (record) openForm(record);
        return;
      }

      const wordButton = event.target.closest("[data-word]");
      if (!wordButton) return;

      const record = contracts.find(
        c => c.id === wordButton.dataset.word
      );

      if (!record) return;

      const status = overlay.querySelector("#msk-word-status");
      wordButton.disabled = true;

      try {
        await downloadWord(record, status);
      } catch (error) {
        status.textContent = "Ошибка: " + error.message;
      } finally {
        wordButton.disabled = false;
      }
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
