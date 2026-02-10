(function () {
  function updateQuery(param, value) {
    const url = new URL(window.location.href);
    const params = url.searchParams;

    // сброс страницы, чтобы не попасть на пустую страницу пагинации
    params.delete("p");

    if (!value) {
      params.delete(param);
    } else {
      params.set(param, value);
    }

    window.location.href = url.pathname + "?" + params.toString();
  }

  function getCurrent(param) {
    const url = new URL(window.location.href);
    return url.searchParams.get(param) || "";
  }

  function injectServiceFilter() {
    // Колонка "Услуги" в твоём list_display — это метод services()
    // Django даёт ей класс th.column-services
    const th = document.querySelector("th.column-services");
    if (!th) return;

    // чтобы не вставить два раза при автообновлении
    if (th.querySelector(".service-filter-select")) return;

    const wrap = document.createElement("div");
    wrap.style.marginTop = "6px";

    const select = document.createElement("select");
    select.className = "service-filter-select";
    select.style.width = "100%";

    const options = [
      ["", "Все"],
      ["cleaning_only", "Только клининг"],
      ["merch_only", "Только выкладка"],
      ["both", "Клининг и выкладка"],
    ];

    const current = getCurrent("service");

    for (const [val, label] of options) {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = label;
      if (val === current) opt.selected = true;
      select.appendChild(opt);
    }

    select.addEventListener("change", function () {
      updateQuery("service", this.value);
    });

    wrap.appendChild(select);
    th.appendChild(wrap);
  }

  // Django admin может перерисовывать таблицу при некоторых действиях — подстрахуемся
  document.addEventListener("DOMContentLoaded", function () {
    injectServiceFilter();
    setTimeout(initEmployeeAutocompletePatching, 50);
  });
const $ = django && django.jQuery ? django.jQuery : null;

  function getStoreIdFromRow(selectEl) {
    if (!$) return "";
    const $tr = $(selectEl).closest("tr");
    // Django changelist editable кладёт скрытый id: form-0-id, form-1-id...
    const $idInput = $tr.find('input[type="hidden"][name$="-id"]');
    return $idInput.length ? $idInput.val() : "";
  }

  function patchEmployeeSelect2(selectEl) {
    if (!$) return;

    const $el = $(selectEl);

    // уже пропатчено
    if ($el.data("patchedStoreId")) return;

    const s2 = $el.data("select2");
    if (!s2) return; // select2 ещё не инициализировался

    const opts = s2.options.options;
    if (!opts.ajax) return;

    const oldData = opts.ajax.data;

    opts.ajax.data = function (params) {
      const data = (typeof oldData === "function") ? oldData(params) : (oldData || {});
      data.store_id = getStoreIdFromRow(selectEl);
      return data;
    };

    $el.data("patchedStoreId", true);
  }

  function initEmployeeAutocompletePatching() {
    if (!$) return;

    // поля assigned_employee в списке магазинов
    const selects = document.querySelectorAll('select[name$="assigned_employee"].admin-autocomplete');
    selects.forEach((el) => patchEmployeeSelect2(el));

    // на случай переинициализации — патчим перед открытием
    $(document).on(
      "select2:opening",
      'select[name$="assigned_employee"].admin-autocomplete',
      function () {
        patchEmployeeSelect2(this);
      }
    );
  }
  });
})();
