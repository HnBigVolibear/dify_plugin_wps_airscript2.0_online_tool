from collections.abc import Generator
from typing import Any
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

def parse_json_array(json_str: str, param_name: str) -> list:
    """
    解析JSON数组字符串
    
    Args:
        json_str: JSON数组字符串，如 "[1,2,3]"
        param_name: 参数名称，用于错误提示
    
    Returns:
        解析后的列表
    
    Raises:
        Exception: 解析失败时抛出异常
    """
    if not json_str or json_str.strip() == "":
        return []
    
    try:
        result = json.loads(json_str)
        if not isinstance(result, list):
            raise Exception(f"{param_name}必须是JSON数组格式，如：[1,2,3]")
        return result
    except json.JSONDecodeError:
        raise Exception(f"{param_name}格式错误，必须是有效的JSON数组格式，如：[1,2,3]")

class test_create_pivot(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        source_sheet_name = tool_parameters.get("source_sheet_name", "")
        source_range = tool_parameters.get("source_range", "").strip()
        row_column_indices_str = tool_parameters.get("row_column_indices", "[]").strip()
        column_column_indices_str = tool_parameters.get("column_column_indices", "[]").strip()
        value_column_indices_str = tool_parameters.get("value_column_indices", "[]").strip()
        function_type = tool_parameters.get("function_type", "count").strip()
        target_sheet_name = tool_parameters.get("target_sheet_name", "")
        target_cell = tool_parameters.get("target_cell", "").strip()

        if not source_sheet_name:
            raise Exception("请填入源数据表名称")
        if not source_range:
            raise Exception("请填入源数据区域，如：A1:D100")
        if not value_column_indices_str:
            raise Exception("请填入作为值字段的列索引列表，不能为空")
        if not row_column_indices_str and not column_column_indices_str:
            raise Exception("请至少填入一个行字段或列字段的列索引列表，否则没法做表啊！？")
        if not target_sheet_name:
            raise Exception("请填入透视表放置的工作表名称")
        if not target_cell:
            raise Exception("请填入透视表放置的起始单元格，如：A1")
        
        # 解析JSON数组
        row_column_indices = parse_json_array(row_column_indices_str, "行字段列索引列表")
        column_column_indices = parse_json_array(column_column_indices_str, "列字段列索引列表")
        value_column_indices = parse_json_array(value_column_indices_str, "值字段列索引列表")
        
        # 验证值字段列表不能为空
        if not value_column_indices:
            raise Exception("值字段列索引列表不能为空，请至少提供一个列索引")
        
        # 验证所有索引都是正整数
        for idx_list, name in [(row_column_indices, "行字段"), (column_column_indices, "列字段"), (value_column_indices, "值字段")]:
            for idx in idx_list:
                if not isinstance(idx, int) or idx < 1:
                    raise Exception(f"{name}列索引必须是正整数（从1开始），当前值：{idx}")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.create_pivot(
                source_sheet_name=source_sheet_name,
                source_range=source_range,
                row_column_indices=row_column_indices,
                column_column_indices=column_column_indices,
                value_column_indices=value_column_indices,
                function_type=function_type,
                target_sheet_name=target_sheet_name,
                target_cell=target_cell
            )
            if result and result[0].get('success'):
                pass
            else:
                raise Exception(f"创建透视表失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"创建透视表失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(str(result[0].get('message', '透视表创建成功'))+"\n")
