import os
import json
from google import genai
from google.genai import types
from dotenv import load_dotenv

load_dotenv()

class AIService:
    def __init__(self):
        self.client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
        self.model = "gemini-2.5-flash"
    
    def _extract_json_from_response(self, text: str) -> dict:
        """Extract JSON from response, handling markdown code blocks"""
        # Remove markdown code blocks if present
        text = text.strip()
        if text.startswith("```json"):
            text = text[7:]  # Remove ```json
        elif text.startswith("```"):
            text = text[3:]  # Remove ```
        if text.endswith("```"):
            text = text[:-3]  # Remove closing ```
        
        text = text.strip()
        
        try:
            return json.loads(text)
        except json.JSONDecodeError as e:
            print(f"JSON Parse Error: {e}")
            print(f"Response text: {text}")
            raise ValueError(f"Failed to parse JSON response: {e}")
        
    async def interpret_query(self, query: str, context: dict) -> dict:
        """Interpret user query and determine Excel action"""
        
        system_prompt = """You are an Excel AI assistant. Analyze the user's request and:
            1. Determine the action type (formula, pivot_table, chart, filter, sort, etc.)
            2. Extract necessary parameters
            3. Provide a clear explanation

            Context includes: selected range, sheet data, existing formulas.

            Action types:
            - "formula": For calculations and formulas
            - "pivot_table": For pivot tables
            - "chart": For graphs, plots, visualizations (use keywords: plot, chart, graph, visualize)
            - "filter": For filtering data
            - "sort": For sorting data
            - "other": For other actions

            For FORMULA actions, you MUST include:
            - "formula": the Excel formula (starting with =)
            - "targetCell": the cell address where the formula should go

            For CHART actions, just set action to "chart" and provide explanation.

            Respond ONLY with valid JSON in this exact format (no markdown, no extra text):
            {
                "action": "formula|pivot_table|chart|filter|sort|other",
                "parameters": {},
                "explanation": "Clear explanation of what will be done"
            }
            """
        
        user_message = f"""
            Query: {query}

            Excel Context:
            - Selected Range: {context.get('selectedRange', 'None')}
            - Sheet Name: {context.get('sheetName', 'Unknown')}
            - Data Sample: {context.get('dataSample', [])}
            - Column Headers: {context.get('headers', [])}

            If the user doesn't specify where to put the formula, use the first empty cell after the selected range or data.
            """
        
        response = self.client.models.generate_content(
            model=self.model,
            contents=[
                types.Content(
                    role="user",
                    parts=[types.Part(text=system_prompt + "\n\n" + user_message)],
                )
            ],
            config=types.GenerateContentConfig(
                temperature=0.1,
                response_mime_type="application/json"
            )
        )
        
        return self._extract_json_from_response(response.text)
    
    async def generate_formula(self, query: str, context: dict) -> str:
        """Generate Excel formula from natural language"""
        
        system_prompt = """You are an Excel formula expert. Generate valid Excel formulas.
            - Use proper Excel function syntax
            - Consider the data context provided
            - Return ONLY the formula, starting with =
            """
        
        user_message = f"""
                Create an Excel formula for: {query}

                Context:
                - Column Headers: {context.get('headers', [])}
                - Data Range: {context.get('selectedRange', 'A1')}
        """
        
        response = self.client.models.generate_content(
            model=self.model,
            contents=[
                types.Content(
                    role="user",
                    parts=[types.Part(text=system_prompt + "\n\n" + user_message)],
                )
            ],
            config=types.GenerateContentConfig(
                temperature=0.1
            )
        )
        
        # Clean up the response
        formula = response.text.strip()
        
        # Remove markdown code blocks if present
        if formula.startswith("```"):
            lines = formula.split("\n")
            formula = "\n".join(lines[1:-1]) if len(lines) > 2 else formula
            formula = formula.strip()
        
        # Ensure formula starts with =
        if not formula.startswith("="):
            formula = "=" + formula
            
        return formula

    async def generate_chart(self, query: str, context: dict) -> dict:
        """Generate chart/graph configuration"""
        
        system_prompt = """You are an Excel chart expert. Generate chart configurations.

                Analyze the data provided and create an appropriate chart configuration.

                IMPORTANT: For dataRange, you must specify the EXACT range of data to chart.
                - If headers are in row 1 and data is in rows 2-10, use "A1:B10" (includes headers)
                - If user has selected a range, use that range
                - Make sure to include all relevant data columns

                Respond ONLY with valid JSON in this exact format (no markdown, no extra text):
                {
                    "chartType": "line|bar|column|pie|area|scatter",
                    "dataRange": "A1:B10",
                    "title": "Chart Title",
                    "xAxis": {
                        "column": "column_name_or_range",
                        "title": "X Axis Title"
                    },
                    "yAxis": {
                        "column": "column_name_or_range", 
                        "title": "Y Axis Title"
                    },
                    "position": "E2"
                }

                Chart Types:
                - line: Line chart (trends over time)
                - bar: Horizontal bar chart
                - column: Vertical bar chart (default)
                - pie: Pie chart (parts of a whole)
                - area: Area chart (cumulative values)
                - scatter: Scatter plot (correlation)

                Rules:
                1. dataRange MUST include the headers if present
                2. Choose chart type based on data structure
                3. For time series data, use line charts
                4. For categorical comparisons, use column/bar charts
                5. For parts-of-whole, use pie charts
                """
        
        # Analyze the context to suggest better range
        selected_range = context.get('selectedRange', 'A1')
        headers = context.get('headers', [])
        data_sample = context.get('dataSample', [])
        row_count = context.get('rowCount', 10)
        column_count = context.get('columnCount', 2)
        
        # Build a smart suggestion for data range
        if selected_range and selected_range != 'None':
            suggested_range = selected_range
        else:
            # Estimate range based on data
            suggested_range = f"A1:{chr(65 + column_count - 1)}{row_count}"
        
        user_message = f"""
            Create a chart for: {query}

            Excel Context:
            - Available Columns: {headers}
            - Selected/Suggested Data Range: {suggested_range}
            - Number of Rows: {row_count}
            - Number of Columns: {column_count}
            - Data Sample (first few rows):
            {data_sample[:5] if data_sample else "No sample available"}

            Analyze the data structure:
            - First row appears to be: {"headers" if headers else "data"}
            - Data type: {"numeric" if any(isinstance(cell, (int, float)) for row in data_sample for cell in row if row) else "mixed"}

            Choose the most appropriate chart type and ensure dataRange captures all relevant data.
            If headers exist, include them in the range (e.g., A1:B10 for headers in row 1, data in rows 2-10).
            """
                
        response = self.client.models.generate_content(
            model=self.model,
            contents=[
                types.Content(
                    role="user",
                    parts=[types.Part(text=system_prompt + "\n\n" + user_message)],
                )
            ],
            config=types.GenerateContentConfig(
                temperature=0.1,
                response_mime_type="application/json"
            )
        )
        
        chart_config = self._extract_json_from_response(response.text)
        
        # Validate and fix data range if needed
        if 'dataRange' not in chart_config or not chart_config['dataRange']:
            chart_config['dataRange'] = suggested_range
        
        return chart_config
    async def generate_pivot_table(self, query: str, context: dict) -> dict:
        """Generate pivot table configuration"""
        
        system_prompt = """You are an Excel pivot table expert. Generate pivot table configurations.

    Analyze the user's request and the data structure to create an appropriate pivot table.

    CRITICAL RULES:
    1. Column names MUST exactly match the headers provided in the context
    2. The "values" array MUST NEVER be empty - always include at least one field to aggregate
    3. If user mentions filtering (e.g., "for Germany", "Midmarket segment"), put those fields in "filters" array
    4. Default aggregation: use "count" for text fields, "sum" for numeric fields

    Field purposes:
    - rows: Fields to group by (categories, dimensions) - what you want to see broken down
    - columns: Fields to spread across columns (optional) - for cross-tabulation
    - values: Fields to aggregate (REQUIRED) - what you want to calculate/measure
    - filters: Fields to filter by - when user says "for X" or "in Y"

    Respond ONLY with valid JSON in this exact format (no markdown):
    {
        "rows": ["exact_column_name"],
        "columns": [],
        "values": [
            {
                "field": "exact_column_name",
                "function": "sum|count|average|max|min"
            }
        ],
        "filters": ["exact_column_name"]
    }

    REMEMBER: "values" array must ALWAYS have at least one item!
    """
        
        headers = context.get('headers', [])
        data_sample = context.get('dataSample', [])
        
        # Analyze data types to suggest numeric columns
        numeric_columns = []
        if data_sample and len(data_sample) > 1:
            for col_idx, header in enumerate(headers):
                try:
                    # Check if most values in this column are numeric
                    sample_values = [row[col_idx] for row in data_sample[1:] if len(row) > col_idx]
                    numeric_count = sum(1 for v in sample_values if isinstance(v, (int, float)) or (isinstance(v, str) and v.replace('.','').replace('-','').isdigit()))
                    if numeric_count > len(sample_values) / 2:
                        numeric_columns.append(header)
                except:
                    pass
        
        user_message = f"""
    Create a pivot table for: {query}

    Available Column Headers (USE THESE EXACTLY): {headers}
    Detected Numeric Columns (good for values): {numeric_columns}
    Data Range: {context.get('selectedRange', 'A1')}

    Data Sample (first 5 rows):
    {data_sample[:5] if data_sample else "No data"}

    Instructions:
    1. Identify FILTER fields: Look for phrases like "for X", "in Y", "where Z" - these go in filters
    2. Identify ROW fields: What categories to group/break down by
    3. Identify VALUE fields: What to measure/aggregate (REQUIRED - never leave empty!)
    - If user doesn't specify, use a numeric column with "sum" or any column with "count"
    - Prefer numeric columns for sum/average
    - Use count for categorical data
    4. Use EXACT column names from headers list

    Examples:
    Query: "pivot table for Midmarket segment in Germany"
    - filters: ["Segment", "Country"] (will need to be filtered to Midmarket and Germany)
    - rows: ["Product"] (or another dimension to analyze)
    - values: [{{"field": "Sales", "function": "sum"}}] (or count if no numeric field)

    Query: "show sales by product"
    - rows: ["Product"]
    - values: [{{"field": "Sales", "function": "sum"}}]

    Query: "count customers by region"
    - rows: ["Region"]
    - values: [{{"field": "Customer", "function": "count"}}]

    CRITICAL: Always include at least one field in values array!
    """
        
        response = self.client.models.generate_content(
            model=self.model,
            contents=[
                types.Content(
                    role="user",
                    parts=[types.Part(text=system_prompt + "\n\n" + user_message)],
                )
            ],
            config=types.GenerateContentConfig(
                temperature=0.1,
                response_mime_type="application/json"
            )
        )
        
        pivot_config = self._extract_json_from_response(response.text)
        
        # Validate that field names exist in headers
        available_headers = [str(h).strip() for h in headers if h]
        
        # Validate and fix configuration
        if 'rows' in pivot_config:
            pivot_config['rows'] = [r for r in pivot_config['rows'] if r in available_headers]
        else:
            pivot_config['rows'] = []
        
        if 'columns' in pivot_config:
            pivot_config['columns'] = [c for c in pivot_config['columns'] if c in available_headers]
        else:
            pivot_config['columns'] = []
        
        if 'values' in pivot_config:
            pivot_config['values'] = [
                v for v in pivot_config['values'] 
                if v.get('field') in available_headers
            ]
        else:
            pivot_config['values'] = []
        
        # CRITICAL FIX: If values is empty, add a default
        if not pivot_config['values']:
            # Try to find a numeric column
            if numeric_columns:
                pivot_config['values'] = [{
                    "field": numeric_columns[0],
                    "function": "sum"
                }]
            elif available_headers:
                # Fall back to counting the first column
                pivot_config['values'] = [{
                    "field": available_headers[0],
                    "function": "count"
                }]
        
        if 'filters' in pivot_config:
            pivot_config['filters'] = [f for f in pivot_config['filters'] if f in available_headers]
        else:
            pivot_config['filters'] = []
        
        return pivot_config