
from collections.abc import Generator
from typing import Any
import io

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_get_used_range_data(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        is_get_data = tool_parameters.get("is_get_data", "是")
        sheet_name = tool_parameters.get("sheet_name", "")
        # 新增参数: 是否返回 Excel 文件 (默认 False)
        return_excel = tool_parameters.get("return_excel", "否") == "是"
        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.get_used_range_data(is_get_data, sheet_name)

            if result:
                if is_get_data == '是':
                    rows_count = len(result)
                    cols_count = len(result[0]) if result and result[0] else 0
                    if rows_count > 0 and cols_count > 0:
                        yield self.create_text_message(f"{str(result)}\n")
                        yield self.create_text_message(f"\n数据获取完毕！已使用区域数据：{rows_count}行 x {cols_count}列。\n注：您可以查看本节点返回的json变量，获取原始数据！\n")
                        yield self.create_json_message({"result": result, "rows_count": rows_count, "cols_count": cols_count})

                        # 如果需要返回 Excel 文件，使用 openpyxl 将数据写入 xlsx 并通过 create_blob_message 返回
                        if return_excel:
                            wb = None
                            try:
                                # optional import; added to requirements when needed
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
                                    # 如果行不是列表/元组，强制转换为单元素列表
                                    row = [row]
                                for c_idx, cell in enumerate(row, start=1):
                                    # openpyxl 可以直接写入 Python 原生类型
                                    ws.cell(row=r_idx, column=c_idx, value=cell) # type: ignore

                            buf = io.BytesIO()
                            wb.save(buf)
                            buf.seek(0)
                            res_file = buf.getvalue()
                            yield self.create_blob_message(res_file, meta={
                                "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "filename": f"Sheet_{sheet_name}.xlsx" if sheet_name else "Sheet1.xlsx",
                            })
                else:
                    yield self.create_text_message(f"\n已使用区域的范围位置如下：\n起始行：{result[0]}\n起始列：{result[1]}\n结束行：{result[2]}\n结束列：{result[3]}\n注：您可以查看本节点返回的json变量，获取原始列表！\n")
                    yield self.create_json_message({ "rowStart": result[0], "colStart": result[1], "rowEnd": result[2], "colEnd": result[3] })
                return
            yield self.create_text_message(f"获取失败！未获取到数据\n{str(result)}\n")
            yield self.create_json_message({"result": result, "rows_count": 0, "cols_count": 0})
        except Exception as e:
            raise Exception(f"获取已使用区域数据失败！WPS官方返回错误信息：{e}")
