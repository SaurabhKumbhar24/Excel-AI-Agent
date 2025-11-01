class ExcelInterpreter:
    """Convert AI responses to Office.js executable code"""
    
    def generate_action(self, ai_response: dict) -> dict:
        action_type = ai_response.get("action")
        
        if action_type == "formula":
            return self._generate_formula_code(ai_response)
        elif action_type == "pivot_table":
            return self._generate_pivot_code(ai_response)
        elif action_type == "chart":
            return self._generate_chart_code(ai_response)
        else:
            return self._generate_generic_code(ai_response)
    
    def _generate_pivot_code(self, ai_response: dict) -> dict:
        params = ai_response.get("parameters", {})
        
        # Validate that we have values
        if not params.get('values'):
            params['values'] = [{"field": params.get('rows', [''])[0], "function": "count"}]
        
        office_js_code = f"""
    await Excel.run(async (context) => {{
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeToAnalyze = sheet.getUsedRange();
        
        // Create pivot table
        const pivotTable = sheet.pivotTables.add(
            "AIPivotTable_" + Date.now(),
            rangeToAnalyze,
            sheet.getRange("A1")
        );
        
        {self._generate_pivot_fields_code(params.get('rows', []), 'row', 'pivotTable')}
        {self._generate_pivot_fields_code(params.get('columns', []), 'column', 'pivotTable')}
        {self._generate_pivot_fields_code(params.get('filters', []), 'filter', 'pivotTable')}
        {self._generate_pivot_values_code(params.get('values', []), 'pivotTable')}
        
        await context.sync();
    }});
    """
        
        return {
            "action": "pivot_table",
            "parameters": params,
            "explanation": ai_response.get("explanation", ""),
            "office_js_code": office_js_code
        }

    def _generate_pivot_fields_code(self, fields: list, axis: str, table_var: str = 'pivotTable') -> str:
        """Generate code for adding fields to pivot table"""
        if not fields:
            return ""
        
        code = ""
        for field in fields:
            hierarchy = f'{table_var}.hierarchies.getItem("{field}")'
            
            if axis == 'row':
                code += f'    {table_var}.rowHierarchies.add({hierarchy});\n'
            elif axis == 'column':
                code += f'    {table_var}.columnHierarchies.add({hierarchy});\n'
            elif axis == 'filter':
                code += f'    {table_var}.filterHierarchies.add({hierarchy});\n'
        
        return code

    def _generate_pivot_values_code(self, values: list, table_var: str = 'pivotTable') -> str:
        """Generate code for adding value fields to pivot table"""
        if not values:
            return ""
        
        code = ""
        for value in values:
            field = value.get('field')
            function = value.get('function', 'sum').lower()
            
            # Map function names to Excel aggregation types
            function_map = {
                'sum': 'Excel.AggregationFunction.sum',
                'count': 'Excel.AggregationFunction.count',
                'average': 'Excel.AggregationFunction.average',
                'max': 'Excel.AggregationFunction.max',
                'min': 'Excel.AggregationFunction.min'
            }
            
            excel_function = function_map.get(function, 'Excel.AggregationFunction.sum')
            
            hierarchy = f'{table_var}.hierarchies.getItem("{field}")'
            code += f'    const dataHierarchy = {table_var}.dataHierarchies.add({hierarchy});\n'
            code += f'    dataHierarchy.summarizeBy = {excel_function};\n'
        
        return code
    
    def _generate_formula_code(self, ai_response: dict) -> dict:
        params = ai_response.get("parameters", {})
        formula = params.get("formula", "")
        
        # Get target cell, default to selected cell or A1
        target_cell = params.get("targetCell", params.get("target", "A1"))
        
        office_js_code = f"""
            await Excel.run(async (context) => {{
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange("{target_cell}");
                range.formulas = [["{formula}"]];
                await context.sync();
            }});
            """
        
        return {
            "action": "formula",
            "parameters": {
                "formula": formula,
                "targetCell": target_cell
            },
            "explanation": ai_response.get("explanation", ""),
            "office_js_code": office_js_code
        }
    
    def _generate_generic_code(self, ai_response: dict) -> dict:
        return {
            "action": "generic",
            "parameters": ai_response.get("parameters", {}),
            "explanation": ai_response.get("explanation", ""),
            "office_js_code": "// Action not yet implemented"
        }
    
    def _generate_chart_code(self, ai_response: dict) -> dict:
        params = ai_response.get("parameters", {})
        
        chart_type_mapping = {
            "line": "Excel.ChartType.line",
            "bar": "Excel.ChartType.barClustered",
            "column": "Excel.ChartType.columnClustered",
            "pie": "Excel.ChartType.pie",
            "area": "Excel.ChartType.area",
            "scatter": "Excel.ChartType.xyscatter"
        }
        
        chart_type = params.get("chartType", "column")
        excel_chart_type = chart_type_mapping.get(chart_type, "Excel.ChartType.columnClustered")
        
        office_js_code = f"""
            await Excel.run(async (context) => {{
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const dataRange = sheet.getRange("{params.get('dataRange', 'A1:B10')}");
                
                const chart = sheet.charts.add(
                    {excel_chart_type},
                    dataRange,
                    Excel.ChartSeriesBy.auto
                );
                
                chart.title.text = "{params.get('title', 'Chart')}";
                chart.legend.position = Excel.ChartLegendPosition.bottom;
                chart.legend.visible = true;
                
                chart.top = 20;
                chart.left = 400;
                chart.height = 300;
                chart.width = 500;
                
                await context.sync();
            }});
            """
        
        return {
            "action": "chart",
            "parameters": params,
            "explanation": ai_response.get("explanation", ""),
            "office_js_code": office_js_code
        }