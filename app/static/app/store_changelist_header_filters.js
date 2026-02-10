(function () {
  function qs(sel, root) { return (root || document).querySelector(sel); }
  function qsa(sel, root) { return Array.from((root || document).querySelectorAll(sel)); }
  function norm(s) { return (s || "").replace(/\s+/g, " ").trim(); }

  function findUlForFilterTitle(root, titleText) {
    const title = titleText;

    // В Django 6+ фильтры могут быть в <details><summary>...</summary><ul>...</ul></details>
    // либо h3 + ul, либо другие контейнеры. Ищем максимально гибко.
    const candidates = qsa("summary, h3, h2, .filter-title, .changelist-filter-title, *", root)
      .filter(el => norm(el.textContent) === title);

    for (const el of candidates) {
      // 1) ul рядом
      let ul = el.nextElementSibling && el.nextElementSibling.tagName === "UL"
        ? el.nextElementSibling
        : null;
      if (ul) return ul;

      // 2) ul внутри родителя (например <details> или <div>)
      const parent = el.closest("details, section, div, fieldset, li") || el.parentElement;
      if (parent) {
        ul = parent.querySelector("ul");
        if (ul) return ul;
      }

      // 3) ul внутри самого элемента (на всякий)
      ul = el.querySelector && el.querySelector("ul");
      if (ul) return ul;
    }

    return null;
  }

  function activeLabel(ul) {
    const a = ul.querySelector("li.selected a");
    return a ? norm(a.textContent) : "Все";
  }

  function makeDropdown(title, ul) {
    const wrap = document.createElement("span");
    wrap.className = "top-filter";

    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "top-filter__btn";
    btn.textContent = `${title}: ${activeLabel(ul)}`;

    const menu = document.createElement("div");
    menu.className = "top-filter__menu";
    menu.hidden = true;

    const list = document.createElement("div");
    list.className = "top-filter__list";

    qsa("a", ul).forEach(a => {
      const link = a.cloneNode(true);
      link.classList.add("top-filter__item");
      list.appendChild(link);
    });

    menu.appendChild(list);

    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      menu.hidden = !menu.hidden;
    });

    document.addEventListener("click", () => { menu.hidden = true; });

    wrap.appendChild(btn);
    wrap.appendChild(menu);
    return wrap;
  }

  function mountTopFilters() {
    const filterRoot = qs("#changelist-filter");
    if (!filterRoot) return;

    // Куда вставлять: сразу под #toolbar (строка поиска)
    const toolbar = qs("#toolbar");
    const anchor = toolbar ? toolbar : (qs("#changelist .actions") || qs("#changelist"));
    if (!anchor) return;

    if (qs(".top-filters")) return; // не дублируем

    const statusUl = findUlForFilterTitle(filterRoot, "Статус");
    const serviceUl = findUlForFilterTitle(filterRoot, "Услуги");

    if (!statusUl && !serviceUl) return;

    const panel = document.createElement("div");
    panel.className = "top-filters";

    if (statusUl) panel.appendChild(makeDropdown("Статус", statusUl));
    if (serviceUl) panel.appendChild(makeDropdown("Услуги", serviceUl));

    // вставляем после toolbar (или перед actions)
    if (toolbar) {
      toolbar.parentNode.insertBefore(panel, toolbar.nextSibling);
    } else {
      anchor.parentNode.insertBefore(panel, anchor);
    }
  }

  document.addEventListener("DOMContentLoaded", function () {
    mountTopFilters();
    setTimeout(mountTopFilters, 200);
    setTimeout(mountTopFilters, 800);
  });
})();
