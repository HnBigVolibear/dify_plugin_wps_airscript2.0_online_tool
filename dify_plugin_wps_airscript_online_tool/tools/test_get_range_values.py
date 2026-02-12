
from collections.abc import Generator
from typing import Any
import io

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient


class test_get_range_values(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")
        # 新增参数: 是否返回 Excel 文件 (默认 False)
        return_excel = tool_parameters.get("return_excel", "否") == "是"

        if not address or address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的区域地址，例如：A1:D3")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            resp = client.get_range_values(address, sheet_name)
            # resp example: [{ 'success': True, 'values': [[...], [...]] }]
            if resp and resp[0].get('success'):
                pass
            else:
                raise Exception(f"获取区域值失败！WPS官方返回错误信息：{resp}")
        except Exception as e:
            raise Exception(f"获取区域值失败！WPS官方返回错误信息：{e}")

        result = resp[0].get('values', [])
        # 先返回 json 与文本（与已有行为一致）
        yield self.create_json_message({"result": result})
        if not result:
            display_text = "该区域为空！！！"
        else:
            display_text = str(result)
        yield self.create_text_message(display_text + "\n")

        # 如果需要返回 Excel 文件，使用 openpyxl 将数据写入 xlsx 并通过 create_blob_message 返回
        if return_excel:
            wb = None
            try:
                # import only when needed, mirror pattern in test_get_used_range_data
                from openpyxl import Workbook
                wb = Workbook()
            except Exception:
                pass
            if wb is None:
                raise Exception("openpyxl 未安装，无法生成 Excel 文件，请联系Dify平台管理员 在 requirements.txt 中添加 openpyxl 并安装依赖")

            ws = wb.active
            ws.title = sheet_name if sheet_name else "Sheet1" # type: ignore
            for r_idx, row in enumerate(result, start=1):
                if not isinstance(row, (list, tuple)):
                    row = [row]
                for c_idx, cell in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=cell) # type: ignore

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            res_file = buf.getvalue()
            yield self.create_blob_message(res_file, meta={
                "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "filename": f"Sheet_{sheet_name}.xlsx" if sheet_name else "Sheet1.xlsx",
            })
