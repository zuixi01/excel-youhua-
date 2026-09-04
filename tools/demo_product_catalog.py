from __future__ import annotations

import json
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse


CATEGORIES = [
    {
        "id": "cat-phone",
        "name": "手机",
        "parent_id": "digital",
        "path": ["数码", "手机"],
        "aliases": ["智能手机", "手机数码"],
        "active": True,
    },
    {
        "id": "cat-shoe",
        "name": "运动鞋",
        "parent_id": "fashion",
        "path": ["服饰", "鞋靴", "运动鞋"],
        "aliases": ["跑鞋", "休闲运动鞋"],
        "active": True,
    },
]

ATTRIBUTES = {
    "cat-phone": [
        {"id": "brand", "title": "品牌", "aliases": ["品牌名称"], "type": "string", "required": True, "multiple": False, "display_order": 10},
        {"id": "model", "title": "型号", "aliases": ["产品型号"], "type": "string", "required": True, "multiple": False, "display_order": 20},
        {"id": "price", "title": "销售价", "aliases": ["价格", "售价"], "type": "decimal", "required": True, "multiple": False, "display_order": 30, "number_format": "0.00"},
    ],
    "cat-shoe": [
        {"id": "brand", "title": "品牌", "aliases": ["品牌名称"], "type": "string", "required": True, "multiple": False, "display_order": 10},
        {"id": "material", "title": "鞋面材质", "aliases": ["材质"], "type": "string", "required": False, "multiple": False, "display_order": 20},
        {"id": "price", "title": "销售价", "aliases": ["价格", "售价"], "type": "decimal", "required": True, "multiple": False, "display_order": 30, "number_format": "0.00"},
    ],
}

SPECIFICATIONS = {
    "cat-phone": [
        {"id": "color", "title": "颜色", "aliases": ["机身颜色"], "type": "enum", "required": True, "multiple": True, "display_order": 10, "enum_values": ["黑色", "白色", "蓝色"]},
        {"id": "storage", "title": "存储容量", "aliases": ["内存", "容量"], "type": "enum", "required": True, "multiple": True, "display_order": 20, "enum_values": ["128GB", "256GB", "512GB"]},
    ],
    "cat-shoe": [
        {"id": "color", "title": "颜色", "aliases": ["鞋身颜色"], "type": "enum", "required": True, "multiple": True, "display_order": 10, "enum_values": ["黑色", "白色", "红色"]},
        {"id": "size", "title": "尺码", "aliases": ["鞋码"], "type": "enum", "required": True, "multiple": True, "display_order": 20, "enum_values": ["38", "39", "40", "41", "42"]},
    ],
}


class CatalogHandler(BaseHTTPRequestHandler):
    server_version = "ExcelAuditorDemoCatalog/1.0"

    def do_GET(self) -> None:  # noqa: N802 - BaseHTTPRequestHandler contract
        request = urlparse(self.path)
        path = request.path.rstrip("/")
        if path == "/health":
            self._json({"status": "ok"})
            return
        if path == "/api/v1/categories":
            query = parse_qs(request.query)
            page = max(int(query.get("page", ["1"])[0]), 1)
            size = max(int(query.get("size", ["500"])[0]), 1)
            start = (page - 1) * size
            self._json({"data": CATEGORIES[start : start + size], "total": len(CATEGORIES)})
            return
        parts = path.split("/")
        if len(parts) == 6 and parts[1:4] == ["api", "v1", "categories"]:
            category_id, resource = parts[4], parts[5]
            collection = ATTRIBUTES if resource == "attributes" else SPECIFICATIONS if resource == "specifications" else None
            if collection is not None and category_id in collection:
                self._json({"data": collection[category_id]})
                return
        self._json({"error": "not_found"}, status=HTTPStatus.NOT_FOUND)

    def log_message(self, format: str, *args: object) -> None:
        print(f"catalog {self.address_string()} {format % args}", flush=True)

    def _json(self, payload: object, *, status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def run() -> None:
    server = ThreadingHTTPServer(("127.0.0.1", 8011), CatalogHandler)
    print("Demo product catalog: http://127.0.0.1:8011", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    run()
