
(() => {
  "use strict";

  const STORAGE_KEY = "msk_demo_contracts_v1";

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

  function findButton(text) {
    return [...document.querySelectorAll("button, a")]
      .find(el => el.textContent.trim().toLowerCase()
        .includes(text.toLowerCase()));
  }

  function showMessage(message) {
    alert(message);
  }

  function createContract() {
    const company = prompt("Введите название контрагента:");
    if (company === null) return;
    if (!company.trim()) {
      showMessage("Название организации не может быть пустым.");
      return;
    }

    const inn = prompt("Введите ИНН контрагента:", "");
    if (inn === null) return;

    const contracts = loadContracts();

    const contract = {
      id: crypto.randomUUID(),
      company: company.trim(),
      inn: inn.trim(),
      createdAt: new Date().toISOString(),
      status: "Черновик"
    };

    contracts.push(contract);

    try {
      saveContracts(contracts);
      showMessage(
        "Черновик сохранён в этом браузере!\n\n" +
        "Контрагент: " + contract.company + "\n" +
        "ИНН: " + contract.inn + "\n\n" +
        "Это демонстрационный черновик, не официальный договор."
      );
      updateCounter();
    } catch (error) {
      showMessage("Не удалось сохранить черновик: " + error.message);
    }
  }

  function showRegistry() {
    const contracts = loadContracts();

    if (!contracts.length) {
      showMessage("Реестр черновиков пока пуст.");
      return;
    }

    const rows = contracts.map((c, index) =>
      (index + 1) + ". " + c.company +
      " | ИНН: " + (c.inn || "не указан") +
      " | " + new Date(c.createdAt).toLocaleDateString("ru-RU")
    );

    showMessage(
      "ЛОКАЛЬНЫЙ РЕЕСТР ЧЕРНОВИКОВ\n\n" +
      rows.join("\n") +
      "\n\nЭто не общий реестр MSK."
    );
  }

  function updateCounter() {
    const count = loadContracts().length;
    const labels = [...document.querySelectorAll("*")]
      .filter(el => el.children.length === 0 &&
        el.textContent.trim() === "Черновики в браузере");

    labels.forEach(label => {
      const card = label.parentElement;
      if (!card) return;
      const elements = [...card.querySelectorAll("*")];
      const value = elements.find(el =>
        el.children.length === 0 &&
        /^\d+$/.test(el.textContent.trim())
      );
      if (value) value.textContent = String(count);
    });
  }

  document.addEventListener("click", event => {
    const button = event.target.closest("button, a");
    if (!button) return;

    const label = button.textContent.trim().toLowerCase();

    if (label.includes("создать договор")) {
      event.preventDefault();
      createContract();
    } else if (label.includes("реестр договоров")) {
      event.preventDefault();
      showRegistry();
    }
  });

  document.addEventListener("DOMContentLoaded", updateCounter);
})();
