"""
Seed all camera products into the database.
Run once: python seed.py
Or called automatically from app.py on first startup.
"""
import json

# Pricing columns:
#   price    = day 1 (VND)
#   price_d2 = day 2 individual rate  (= "2 NGÀY total" - "1 NGÀY")
#   price_d3 = day 3 individual rate  (= ">3 NGÀY cộng thêm")
#   price_d4 = day 4+ rate per day    (same as price_d3)

SEED_CAMERAS = [
    # ── Canon ──────────────────────────────────────────────────────────────
    {
        'name': 'Canon IXY 2000IS',
        'slug': 'canon-ixy-2000is',
        'brand': 'Canon',
        'type': 'Rent',
        'price': 130000, 'price_d2': 110000, 'price_d3': 70000, 'price_d4': 70000,
        'import_cost': 1200000,
        'featured': False, 'badge': '', 'stock': 2,
        'description': 'Canon IXY 2000IS – Máy compact nhỏ gọn, chụp ảnh đẹp cho du lịch thường ngày.',
        'specs': {
            'Cảm biến': '12.1 MP 1/2.3"',
            'Zoom quang học': '4x',
            'Video': '1080p@30fps',
        },
    },
    {
        'name': 'Canon EOS M10',
        'slug': 'canon-eos-m10',
        'brand': 'Canon',
        'type': 'Rent',
        'price': 150000, 'price_d2': 120000, 'price_d3': 90000, 'price_d4': 90000,
        'import_cost': 2500000,
        'featured': True, 'badge': '', 'stock': 2,
        'description': 'Canon EOS M10 – Mirrorless giá tốt, nhỏ gọn, cảm biến 18MP, màn hình lật selfie 180°, phù hợp cho chụp du lịch hàng ngày.',
        'specs': {
            'Cảm biến': '18 MP APS-C CMOS',
            'Màn hình': '3" lật 180° cảm ứng',
            'Video': '1080p@30fps',
            'Kết nối': 'Wi-Fi, NFC',
        },
    },
    {
        'name': 'Canon EOS M3',
        'slug': 'canon-eos-m3',
        'brand': 'Canon',
        'type': 'Rent',
        'price': 160000, 'price_d2': 130000, 'price_d3': 100000, 'price_d4': 100000,
        'import_cost': 3000000,
        'featured': False, 'badge': '', 'stock': 2,
        'description': 'Canon EOS M3 – Mirrorless APS-C 24.2MP với kính ngắm EVF tùy chọn, màn hình lật xoay, Wi-Fi, lý tưởng cho nhiếp ảnh du lịch.',
        'specs': {
            'Cảm biến': '24.2 MP APS-C CMOS',
            'Màn hình': '3" lật xoay cảm ứng',
            'Video': '1080p@30fps',
            'Kết nối': 'Wi-Fi, NFC',
        },
    },
    {
        'name': 'Canon EOS M100',
        'slug': 'canon-eos-m100',
        'brand': 'Canon',
        'type': 'Rent',
        'price': 170000, 'price_d2': 140000, 'price_d3': 100000, 'price_d4': 100000,
        'import_cost': 3500000,
        'featured': True, 'badge': '', 'stock': 2,
        'description': 'Canon EOS M100 – Mirrorless compact thời trang, cảm biến 24.2MP APS-C, màn hình lật selfie 180°, Wi-Fi/Bluetooth.',
        'specs': {
            'Cảm biến': '24.2 MP APS-C CMOS',
            'Màn hình': '3" lật 180° cảm ứng',
            'Video': '1080p@60fps',
            'Kết nối': 'Wi-Fi, Bluetooth, NFC',
        },
    },
    {
        'name': 'Canon EOS M6',
        'slug': 'canon-eos-m6',
        'brand': 'Canon',
        'type': 'Rent',
        'price': 170000, 'price_d2': 140000, 'price_d3': 110000, 'price_d4': 110000,
        'import_cost': 4500000,
        'featured': True, 'badge': 'Hot', 'stock': 2,
        'description': 'Canon EOS M6 – Mirrorless APS-C 24.2MP thiết kế rangefinder đẹp, tốc độ chụp nhanh, AF Dual Pixel, màn hình lật xoay.',
        'specs': {
            'Cảm biến': '24.2 MP APS-C CMOS',
            'AF': 'Dual Pixel CMOS AF',
            'Tốc độ chụp': '9 fps',
            'Màn hình': '3" lật xoay cảm ứng',
            'Video': '1080p@60fps',
            'Kết nối': 'Wi-Fi, Bluetooth, NFC',
        },
    },

    # ── Fujifilm ────────────────────────────────────────────────────────────
    {
        'name': 'Fujifilm X-T10',
        'slug': 'fujifilm-x-t10',
        'brand': 'Fujifilm',
        'type': 'Rent',
        'price': 200000, 'price_d2': 170000, 'price_d3': 140000, 'price_d4': 140000,
        'import_cost': 5000000,
        'featured': True, 'badge': '', 'stock': 2,
        'description': 'Fujifilm X-T10 – Mirrorless APS-C phong cách retro, cảm biến 16.3MP X-Trans II, Film Simulation màu cổ điển đẹp, gọn nhẹ.',
        'specs': {
            'Cảm biến': '16.3 MP X-Trans CMOS II (APS-C)',
            'Tốc độ chụp': '8 fps',
            'Video': '1080p@60fps',
            'Film Simulation': '15 chế độ',
            'Kết nối': 'Wi-Fi',
        },
    },
    {
        'name': 'Fujifilm X-T20',
        'slug': 'fujifilm-x-t20',
        'brand': 'Fujifilm',
        'type': 'Rent',
        'price': 230000, 'price_d2': 170000, 'price_d3': 170000, 'price_d4': 170000,
        'import_cost': 7000000,
        'featured': True, 'badge': '', 'stock': 2,
        'description': 'Fujifilm X-T20 – Mirrorless APS-C retro 24.3MP, Film Simulation phong phú, màn hình lật 3 chiều, quay 4K. Lựa chọn tốt cho người yêu màu Fuji.',
        'specs': {
            'Cảm biến': '24.3 MP X-Trans CMOS III (APS-C)',
            'Tốc độ chụp': '14 fps',
            'Video': '4K15p / 1080p@60fps',
            'Màn hình': '3" lật 3 chiều cảm ứng',
            'Film Simulation': '19 chế độ',
        },
    },

    # ── Sony ────────────────────────────────────────────────────────────────
    {
        'name': 'Sony NEX-3N',
        'slug': 'sony-nex-3n',
        'brand': 'Sony',
        'type': 'Rent',
        'price': 140000, 'price_d2': 100000, 'price_d3': 80000, 'price_d4': 80000,
        'import_cost': 1800000,
        'featured': False, 'badge': '', 'stock': 2,
        'description': 'Sony NEX-3N – Mirrorless APS-C nhỏ gọn với zoom tích hợp, màn hình lật 180° selfie, AF nhanh, màu ảnh Sony đẹp tự nhiên.',
        'specs': {
            'Cảm biến': '16.1 MP APS-C Exmor CMOS',
            'Màn hình': '3" lật 180°',
            'Video': '1080p@60fps',
            'Kết nối': 'Wi-Fi (với NFC)',
        },
    },
    {
        'name': 'Sony A5000',
        'slug': 'sony-a5000',
        'brand': 'Sony',
        'type': 'Rent',
        'price': 150000, 'price_d2': 120000, 'price_d3': 90000, 'price_d4': 90000,
        'import_cost': 2200000,
        'featured': False, 'badge': '', 'stock': 2,
        'description': 'Sony A5000 – Mirrorless APS-C siêu nhẹ (210g), 20.1MP, màn hình selfie 180°, Wi-Fi/NFC, lý tưởng cho chụp du lịch nhẹ nhàng.',
        'specs': {
            'Cảm biến': '20.1 MP APS-C Exmor CMOS',
            'Trọng lượng': '~210g',
            'Màn hình': '3" lật 180°',
            'Video': '1080p@30fps',
            'Kết nối': 'Wi-Fi, NFC',
        },
    },
]


def seed(db, Camera):
    """Insert any products not yet in the DB (idempotent)."""
    added = 0
    for data in SEED_CAMERAS:
        if Camera.query.filter_by(slug=data['slug']).first():
            continue
        data  = dict(data)               # don't mutate the module-level list
        specs = data.pop('specs', {})
        cam = Camera(**data, specs_json=json.dumps(specs, ensure_ascii=False))
        db.session.add(cam)
        added += 1
    if added:
        db.session.commit()
    return added


def backfill_costs(db, Camera):
    """Set import_cost on existing rows that don't have one yet (idempotent)."""
    costs = {c['slug']: c.get('import_cost', 0) for c in SEED_CAMERAS}
    updated = 0
    # Sale cameras only — rental gear acquisition cost is intentionally kept at 0.
    for cam in Camera.query.filter_by(type='Sale').all():
        if (not cam.import_cost) and costs.get(cam.slug):
            cam.import_cost = costs[cam.slug]
            updated += 1
    if updated:
        db.session.commit()
    return updated
