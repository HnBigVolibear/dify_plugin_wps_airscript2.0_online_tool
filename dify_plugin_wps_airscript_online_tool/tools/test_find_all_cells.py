
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_find_all_cells(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        search_text = tool_parameters.get("search_text", "").strip()
        search_range = tool_parameters.get("search_range", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")

        if not search_text:
            raise Exception("请输入要查找的文本")

        if not search_range or search_range[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的搜索范围，例如：A1:Z100")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.find_all_cells(search_text, search_range, sheet_name)

            yield self.create_json_message({"result": result})
            if result and result[0]:
                yield self.create_text_message(f"找到{len(result)}个匹配单元格，详情如下：\n{str(result)}\n")
            else:
                yield self.create_text_message("未找到匹配的单元格\n")
        except Exception as e:
            raise Exception(f"查找所有匹配单元格失败！WPS官方返回错误信息：{e}")
