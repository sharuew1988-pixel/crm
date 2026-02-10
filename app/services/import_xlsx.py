from decimal import Decimal
from django.db import transaction

def process_request_xlsx(request_obj, xlsx_path: str) -> dict:
    df = read_excel_with_header_search(xlsx_path)

    # нормализуем имена колонок
    norm = {str(c).strip().lower(): c for c in df.columns}

    def pick_col(*needles):
        for n in needles:
            n = str(n).strip().lower()
            for k, original in norm.items():
                if n in k:
                    return original
        return None

    col_address = pick_col("адрес объекта")
    col_service = pick_col("вид оказываемых услуг", "услуг")
    col_hours = pick_col("часы", "часов")

    missing = []
    if not col_address:
        missing.append("Адрес объекта")
    if not col_service:
        missing.append("Вид оказываемых услуг")
    if not col_hours:
        missing.append("Часы")
    if missing:
        raise ValueError(f"Не нашёл нужные колонки: {missing}. В файле есть: {list(df.columns)}")

    created_stores = 0
    created_lines = 0

    # суммы часов по магазинам в рамках текущего файла
    store_hours: dict[int, Decimal] = {}

    with transaction.atomic():
        # 1) обнуляем текущие часы у всех магазинов по типу заявки
        if request_obj.source_type == "merch":
            Store.objects.update(current_hours_merch=Decimal("0.00"))
        else:  # cleaning
            Store.objects.update(current_hours_cleaning=Decimal("0.00"))

        # 2) создаём историю и считаем суммы по магазинам для текущего файла
        for _, row in df.iterrows():
            address_raw = str(row[col_address]).strip()
            if not address_raw or address_raw.lower() in ("nan", "none"):
                continue

            service = normalize_service(row[col_service])
            hours = parse_hours(row[col_hours])

            city, address = parse_address(address_raw)

            store, store_created = Store.objects.get_or_create(
                city=city,
                address=address,
                defaults={"address_raw": address_raw},
            )

            if store_created:
                created_stores += 1
            elif store.address_raw != address_raw:
                store.address_raw = address_raw
                store.save(update_fields=["address_raw"])

            # копим часы по магазину (в рамках файла)
            store_hours[store.id] = store_hours.get(store.id, Decimal("0.00")) + hours

            # сохраняем историю строк
            row_hash = make_row_hash(address_raw, service, hours)

            _, line_created = RequestLine.objects.get_or_create(
                request=request_obj,
                row_hash=row_hash,
                defaults={
                    "store": store,
                    "service_type": service,
                    "hours": hours,
                },
            )
            if line_created:
                created_lines += 1

        # 3) проставляем текущие часы только тем магазинам, которые есть в файле
        if store_hours:
            stores_to_update = list(Store.objects.filter(id__in=store_hours.keys()))
            if request_obj.source_type == "merch":
                for s in stores_to_update:
                    s.current_hours_merch = store_hours.get(s.id, Decimal("0.00"))
                Store.objects.bulk_update(stores_to_update, ["current_hours_merch"])
            else:
                for s in stores_to_update:
                    s.current_hours_cleaning = store_hours.get(s.id, Decimal("0.00"))
                Store.objects.bulk_update(stores_to_update, ["current_hours_cleaning"])

    return {"created_stores": created_stores, "created_lines": created_lines}
