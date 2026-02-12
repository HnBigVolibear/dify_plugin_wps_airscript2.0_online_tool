
from collections.abc import Generator
from typing import Any
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_set_range_values(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        values_str = tool_parameters.get("values", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")

        if not address or address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的区域地址，例如：A1:D3")

        if not values_str or values_str[0] != "[" or values_str[-1] != "]":
            raise Exception('请提供要设置的区域内容！必须是有效的JSON格式的二维数组，例如：[["测试1", "测试2"], ["测试3", "测试4]]')

        # 将string转换为数组
        try:
            values = json.loads(values_str)
            if not isinstance(values, list):
                raise Exception("数据格式错误，应为二维数组")
        except json.JSONDecodeError:
            try:
                import ast
                values = ast.literal_eval(values_str)
            except Exception as e:
                raise Exception('数据格式错误，请提供有效的JSON格式的二维数组，例如：[["测试1", "测试2"], ["测试3", "测试4]]')

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.set_range_values(address, values, sheet_name)
            # print(result) # [{'message': '设置成功', 'success': True}]
            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"设置区域值失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"设置区域值失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(f"设置区域值成功！区域：{address}，共写入 {len(values)} 行数据\n")
