(function () {
  document.addEventListener("DOMContentLoaded", function () {
    // если datalist уже есть — не дублируем
    if (document.getElementById("source_suggestions")) return;

    const input = document.querySelector('input[list="source_suggestions"]');
    if (!input) return;

    const datalist = document.createElement("datalist");
    datalist.id = "source_suggestions";

    // Варианты подсказок (должны совпадать с SOURCE_SUGGESTIONS)
    const options = ["HH.ru","Avito","Сайт","VK","Telegram","2ГИС","Рекомендация","Холодный обзвон","Email-рассылка"];
    options.forEach((v) => {
      const o = document.createElement("option");
      o.value = v;
      datalist.appendChild(o);
    });

    document.body.appendChild(datalist);
  });
})();
