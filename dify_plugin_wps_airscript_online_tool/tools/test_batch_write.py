
from collections.abc import Generator
from typing import Any
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_batch_write(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        data_str = tool_parameters.get("data", "").strip()
        start_cell = tool_parameters.get("start_cell", "A1").strip()
        sheet_name = tool_parameters.get("sheet_name", "")

        if not start_cell or start_cell[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的起始单元格，例如：A1")

        if not data_str  or data_str[0] != "[" or data_str[-1] != "]":
            raise Exception('请提供要写入的数据！必须是有效的JSON格式的二维数组，例如：[["测试1", "测试2"], ["测试3", "测试4]]')

        # 将string转换为数组
        try:
            data = json.loads(data_str)
            if not isinstance(data, list):
                raise Exception("数据格式错误，应为二维数组")
        except json.JSONDecodeError:
            try:
                import ast
                data = ast.literal_eval(data_str)
            except Exception as e:
                raise Exception('数据格式错误，请提供有效的JSON格式的二维数组，例如：[["测试1", "测试2"], ["测试3", "测试4]]')

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.batch_write(data, start_cell=start_cell, sheet_name=sheet_name)
            # print(result) # [{'message': '设置成功', 'success': True}]
            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"批量写入失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"批量写入失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(f"批量写入成功！从 {start_cell} 开始写入 {len(data)} 行数据\n")
