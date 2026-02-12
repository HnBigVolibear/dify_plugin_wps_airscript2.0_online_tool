
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_replace_in_range(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        search_text = tool_parameters.get("search_text", "").strip()
        replace_text = tool_parameters.get("replace_text", "")
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
            result = client.replace_in_range(search_text, replace_text, search_range, sheet_name)  
            if result and result[0]['success']:
                result = result[0]['count']
                yield self.create_json_message({"result": result, "info": "替换成功！"})
                yield self.create_text_message(f"替换成功！共替换了{result}处\n")
            else:
                # raise Exception(f"替换内容失败！WPS官方返回错误信息：{result}")
                yield self.create_json_message({"result": 0, "info": "替换失败！"})
                yield self.create_text_message(f"替换失败！共替换了0处\n")
        except Exception as e:
            raise Exception(f"替换内容失败！WPS官方返回错误信息：{e}")

        
