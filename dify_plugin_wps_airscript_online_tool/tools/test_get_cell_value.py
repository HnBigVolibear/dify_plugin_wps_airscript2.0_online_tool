from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_get_cell_value(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")
        
        if not address or address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的单元格，例如：A1")
        if ":" in address:
            raise Exception("请填入单个单元格地址，而非区域地址！例如：A1，而不是错误的比如A1:B1")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')
            
            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.get_cell_value(address, sheet_name)
            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"初始化WPS AirScript接口失败，配置参数鉴权失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"初始化WPS AirScript接口失败，配置参数鉴权失败！WPS官方返回错误信息：{e}")
        
        # print(result)  # [{'success': True, 'value': 1}]
        result = result[0]['value']
        yield self.create_json_message({"result": result})
        if not result:
            result = "该单元格为空！！！"
        yield self.create_text_message(str(result)+"\n")
