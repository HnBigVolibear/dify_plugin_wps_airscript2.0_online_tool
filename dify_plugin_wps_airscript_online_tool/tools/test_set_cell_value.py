from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_get_cell_value(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")
        set_value = tool_parameters.get("set_value", "")

        if not address or address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的单元格，例如：A1")
        if ":" in address:
            yield self.create_text_message("【警告】当前输入的是区域，而非单元格，不过也是兼容的。诶呦，你小子，让你发现盲点了啊？！没错，这是我故意这样设计的。\n")
        
        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')
            
            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.set_cell_value(address, set_value, sheet_name)
            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"初始化WPS AirScript接口失败，配置参数鉴权失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"初始化WPS AirScript接口失败，配置参数鉴权失败！WPS官方返回错误信息：{e}")
        
        # print(result)  # [{'message': '设置成功', 'success': True}]
        result = result[0]['message']
        yield self.create_text_message(str(result)+"\n")
