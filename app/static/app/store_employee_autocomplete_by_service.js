(function () {
  // В Django admin jQuery живёт здесь:
  const $ = window.django && window.django.jQuery ? window.django.jQuery : null;

  if (!$) {
    console.warn("django.jQuery не найден — autocomplete by service отключён");
    return;
  }

  function getStoreIdFromRow($row) {
    // Пытаемся вытащить store_id из ссылки на редактирование (карандаш) в этой строке
    const editLink = $row.find('a[href*="/change/"]').first();
    if (!editLink.length) return null;

    const href = editLink.attr("href"); // например ".../123/change/"
    const m = href.match(/\/(\d+)\/change\/?$/);
    return m ? m[1] : null;
  }

  function patchSelect2Ajax($input) {
    // select2 создаётся на hidden input внутри .admin-autocomplete
    const s2 = $input.data("select2");
    if (!s2 || !s2.options || !s2.options.options || !s2.options.options.ajax) return;

    const oldData = s2.options.options.ajax.data;

    s2.options.options.ajax.data = function (params) {
      const data = oldData ? oldData(params) : params;

      // найдём строку (tr) где находится этот autocomplete
      const $row = $input.closest("tr");
      const storeId = getStoreIdFromRow($row);

      if (storeId) data.store_id = storeId;
      return data;
    };
  }

  function init() {
    // На changelist list_editable поля появляются как autocomplete (select2)
    $(".admin-autocomplete input.select2-search__field").each(function () {
      // это поле поиска внутри dropdown, нам нужно исходное hidden input:
      // но проще — патчить на событии открытия, когда select2 уже есть
    });

    // При открытии любого select2 в админке — патчим его ajax.data
    $(document).on("select2:opening", ".admin-autocomplete", function () {
      const $widget = $(this);

      // исходный input внутри виджета:
      const $input = $widget.find("input.select2-hidden-accessible");
      if ($input.length) patchSelect2Ajax($input);
    });
  }

  $(init);
})();
