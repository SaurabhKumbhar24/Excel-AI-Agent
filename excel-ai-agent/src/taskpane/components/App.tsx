import * as React from "react";
import { useState, useEffect } from "react";
import axios from "axios";

const API_BASE_URL = "http://localhost:8000/api/v1";

interface AIResponse {
  action: string;
  parameters: any;
  explanation: string;
  office_js_code: string;
}

const App: React.FC = () => {
  const [query, setQuery] = useState("");
  const [loading, setLoading] = useState(false);
  const [response, setResponse] = useState<AIResponse | null>(null);
  const [error, setError] = useState("");
  const [successMessage, setSuccessMessage] = useState("");
  const [editedTargetCell, setEditedTargetCell] = useState("");

  // Auto-clear messages after 5 seconds
  useEffect(() => {
    if (successMessage) {
      const timer = setTimeout(() => setSuccessMessage(""), 5000);
      return () => clearTimeout(timer);
    }
    return () => {};
  }, [successMessage]);

  useEffect(() => {
    if (error) {
      const timer = setTimeout(() => setError(""), 5000);
      return () => clearTimeout(timer);
    }
    return () => {};
  }, [error]);

  // Update edited target cell when response changes
  useEffect(() => {
    if (response && response.parameters && response.parameters.targetCell) {
      setEditedTargetCell(response.parameters.targetCell);
    }
  }, [response]);

  const getExcelContext = async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = context.workbook.getSelectedRange();
      
      sheet.load("name");
      range.load("address, values, rowCount, columnCount");
      
      await context.sync();
      
      // Get headers (first row) - ensure they're strings
      const headers = range.values[0].map(h => String(h || '').trim());
      
      // Get all data including headers
      const allData = range.values;
      
      return {
        sheetName: sheet.name,
        selectedRange: range.address.split("!")[1] || range.address, // Remove sheet name if present
        dataSample: allData.slice(0, 10), // First 10 rows including header
        headers: headers,
        rowCount: range.rowCount,
        columnCount: range.columnCount
      };
    });
  };

  const handleQuery = async () => {
    if (!query.trim()) return;
    
    setLoading(true);
    setError("");
    setResponse(null);
    setSuccessMessage("");

    try {
      const context = await getExcelContext();
      
      // First, interpret the query
      const interpretResult = await axios.post(`${API_BASE_URL}/query`, {
        query: query,
        context: context
      });
      
      let finalResult = interpretResult.data;
      
      // If it's a chart action, get detailed chart config (now properly processed)
      if (finalResult.action === "chart") {
        const chartResult = await axios.post(`${API_BASE_URL}/create-chart`, {
          query: query,
          context: context
        });
        
        // chartResult now has the full processed response with office_js_code
        finalResult = chartResult.data;
      }
      
      // If it's a pivot table, get detailed config
      if (finalResult.action === "pivot_table") {
        const pivotResult = await axios.post(`${API_BASE_URL}/create-pivot-table`, {
          query: query,
          context: context
        });
        
        // Update parameters with detailed config
        finalResult.parameters = pivotResult.data;
      }
      
      setResponse(finalResult);
    } catch (err: any) {
      setError(err.response?.data?.detail || "An error occurred");
      console.error("Error:", err);
    } finally {
      setLoading(false);
    }
  };

  const executeAction = async () => {
    if (!response) return;
    
    setError("");
    setSuccessMessage("");
    
    try {
      if (response.action === "formula") {
        await executeFormula(response.parameters);
      } else if (response.action === "pivot_table") {
        await executePivotTable(response.parameters);
      } else if (response.action === "chart") {
        await executeChart(response.parameters);
      }
      setSuccessMessage(`Successfully executed ${response.action}!`);
    } catch (err: any) {
      setError("Error executing action: " + err.message);
    }
  };

  const executeFormula = async (params: any) => {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      let targetCell = editedTargetCell || params.targetCell;
      
      if (!targetCell) {
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("address");
        await context.sync();
        targetCell = selectedRange.address.split("!")[1];
      }
      
      const range = sheet.getRange(targetCell);
      range.formulas = [[params.formula]];
      range.select();
      
      await context.sync();
    });
  };

  const executeChart = async (params: any) => {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      const dataRange = sheet.getRange(params.dataRange || "A1:B10");
      
      const chartTypeMapping: any = {
        "line": Excel.ChartType.line,
        "bar": Excel.ChartType.barClustered,
        "column": Excel.ChartType.columnClustered,
        "pie": Excel.ChartType.pie,
        "area": Excel.ChartType.area,
        "scatter": Excel.ChartType.xyscatter
      };
      
      const chartType = chartTypeMapping[params.chartType] || Excel.ChartType.columnClustered;
      
      const chart = sheet.charts.add(
        chartType,
        dataRange,
        Excel.ChartSeriesBy.auto
      );
      
      chart.title.text = params.title || "Chart";
      chart.legend.position = Excel.ChartLegendPosition.bottom;
      chart.legend.visible = true;
      
      if (params.xAxis && params.xAxis.title) {
        chart.axes.categoryAxis.title.text = params.xAxis.title;
      }
      
      if (params.yAxis && params.yAxis.title) {
        chart.axes.valueAxis.title.text = params.yAxis.title;
      }
      
      const position = params.position || "E2";
      try {
        const positionRange = sheet.getRange(position);
        positionRange.load("top, left");
        await context.sync();
        chart.top = positionRange.top;
        chart.left = positionRange.left;
      } catch {
        chart.top = 20;
        chart.left = 400;
      }
      
      chart.height = 300;
      chart.width = 500;
      
      await context.sync();
    });
  };
    
  const executePivotTable = async (params: any) => {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rangeToAnalyze = sheet.getUsedRange();
      
      // Create a new sheet for pivot table to avoid conflicts
      const pivotSheet = context.workbook.worksheets.add("Pivot_" + Date.now());
      
      // Create pivot table
      const pivotTable = pivotSheet.pivotTables.add(
        "AIPivotTable_" + Date.now(),
        rangeToAnalyze,
        pivotSheet.getRange("A3")
      );
      
      // Add filter fields first (so users can filter)
      if (params.filters && params.filters.length > 0) {
        for (const filterField of params.filters) {
          try {
            const hierarchy = pivotTable.hierarchies.getItem(filterField);
            pivotTable.filterHierarchies.add(hierarchy);
          } catch (err) {
            console.warn(`Could not add filter field: ${filterField}`, err);
          }
        }
      }
      
      // Add row fields
      if (params.rows && params.rows.length > 0) {
        for (const rowField of params.rows) {
          try {
            const hierarchy = pivotTable.hierarchies.getItem(rowField);
            pivotTable.rowHierarchies.add(hierarchy);
          } catch (err) {
            console.warn(`Could not add row field: ${rowField}`, err);
          }
        }
      }
      
      // Add column fields
      if (params.columns && params.columns.length > 0) {
        for (const colField of params.columns) {
          try {
            const hierarchy = pivotTable.hierarchies.getItem(colField);
            pivotTable.columnHierarchies.add(hierarchy);
          } catch (err) {
            console.warn(`Could not add column field: ${colField}`, err);
          }
        }
      }
      
      // Add value fields (REQUIRED)
      if (params.values && params.values.length > 0) {
        for (const valueField of params.values) {
          try {
            const hierarchy = pivotTable.hierarchies.getItem(valueField.field);
            const dataHierarchy = pivotTable.dataHierarchies.add(hierarchy);
            dataHierarchy.summarizeBy = getSummarizeBy(valueField.function);
          } catch (err) {
            console.warn(`Could not add value field: ${valueField.field}`, err);
          }
        }
      } else {
        throw new Error("No value fields specified for pivot table");
      }
      
      // Activate the new sheet
      pivotSheet.activate();
      
      await context.sync();
    });
  };

  const getSummarizeBy = (func: string) => {
    const mapping: any = {
      sum: Excel.AggregationFunction.sum,
      count: Excel.AggregationFunction.count,
      average: Excel.AggregationFunction.average,
      max: Excel.AggregationFunction.max,
      min: Excel.AggregationFunction.min
    };
    return mapping[func.toLowerCase()] || Excel.AggregationFunction.sum;
  };

  return (
    <div style={{ padding: "20px", fontFamily: "Segoe UI, sans-serif" }}>
      <h2 style={{ marginTop: 0 }}>Excel AI Agent</h2>
      
      {successMessage && (
        <div style={{
          marginBottom: "15px",
          padding: "12px",
          backgroundColor: "#d4edda",
          border: "1px solid #c3e6cb",
          borderRadius: "4px",
          color: "#155724"
        }}>
          ✓ {successMessage}
        </div>
      )}

      {error && (
        <div style={{
          marginBottom: "15px",
          padding: "12px",
          backgroundColor: "#f8d7da",
          border: "1px solid #f5c6cb",
          borderRadius: "4px",
          color: "#721c24"
        }}>
          ✗ {error}
        </div>
      )}
      
      <div style={{ marginBottom: "20px" }}>
        <textarea
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Ask me to create formulas, pivot tables, charts, or analyze data...&#10;&#10;Examples:&#10;• Create a formula to sum A1 to A10&#10;• Make a pivot table showing sales by region&#10;• Plot a line chart of my data&#10;• Calculate the average of column B"
          style={{
            width: "100%",
            height: "120px",
            padding: "10px",
            fontSize: "14px",
            border: "1px solid #ccc",
            borderRadius: "4px",
            resize: "vertical",
            fontFamily: "inherit"
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter" && e.ctrlKey) {
              handleQuery();
            }
          }}
        />
        <small style={{ color: "#666", fontSize: "12px" }}>
          Press Ctrl+Enter to submit
        </small>
      </div>

      <button
        onClick={handleQuery}
        disabled={loading || !query.trim()}
        style={{
          padding: "12px 20px",
          backgroundColor: loading || !query.trim() ? "#ccc" : "#0078d4",
          color: "white",
          border: "none",
          cursor: loading || !query.trim() ? "not-allowed" : "pointer",
          width: "100%",
          fontSize: "15px",
          borderRadius: "4px",
          fontWeight: "600"
        }}
      >
        {loading ? "Processing..." : "Ask AI"}
      </button>

      {response && (
        <div style={{ 
          marginTop: "20px", 
          border: "1px solid #ddd", 
          padding: "15px", 
          borderRadius: "4px",
          backgroundColor: "#f9f9f9"
        }}>
          <h3 style={{ marginTop: 0 }}>AI Response</h3>
          <p>
            <strong>Action:</strong>{" "}
            <span style={{
              backgroundColor: "#e3f2fd",
              padding: "2px 8px",
              borderRadius: "3px",
              fontSize: "13px"
            }}>
              {response.action}
            </span>
          </p>
          <p><strong>Explanation:</strong> {response.explanation}</p>
          
          {response.action === "formula" && response.parameters && (
            <div style={{ 
              backgroundColor: "#fff", 
              padding: "12px", 
              marginTop: "10px",
              borderRadius: "4px",
              border: "1px solid #e0e0e0"
            }}>
              <p style={{ margin: "5px 0" }}>
                <strong>Formula:</strong>{" "}
                <code style={{
                  backgroundColor: "#f5f5f5",
                  padding: "2px 6px",
                  borderRadius: "3px",
                  fontSize: "13px"
                }}>
                  {response.parameters.formula}
                </code>
              </p>
              <div style={{ marginTop: "10px" }}>
                <label style={{ display: "block", marginBottom: "5px" }}>
                  <strong>Target Cell:</strong>
                </label>
                <input
                  type="text"
                  value={editedTargetCell}
                  onChange={(e) => setEditedTargetCell(e.target.value.toUpperCase())}
                  placeholder="e.g., D1"
                  style={{
                    padding: "6px 10px",
                    width: "100px",
                    border: "1px solid #ccc",
                    borderRadius: "3px",
                    fontSize: "14px"
                  }}
                />
              </div>
            </div>
          )}
          
          {response.action === "pivot_table" && response.parameters && (
            <div style={{ 
              backgroundColor: "#fff", 
              padding: "12px", 
              marginTop: "10px",
              borderRadius: "4px",
              border: "1px solid #e0e0e0"
            }}>
              {response.parameters.filters && response.parameters.filters.length > 0 && (
                <p style={{ margin: "5px 0" }}>
                  <strong>Filters:</strong>{" "}
                  <span style={{
                    backgroundColor: "#fff3cd",
                    padding: "2px 8px",
                    borderRadius: "3px",
                    fontSize: "13px"
                  }}>
                    {JSON.stringify(response.parameters.filters)}
                  </span>
                  <br />
                  <small style={{ color: "#666" }}>
                    You'll need to manually filter these in the pivot table
                  </small>
                </p>
              )}
              <p style={{ margin: "5px 0" }}>
                <strong>Rows:</strong> {JSON.stringify(response.parameters.rows)}
              </p>
              <p style={{ margin: "5px 0" }}>
                <strong>Columns:</strong> {JSON.stringify(response.parameters.columns)}
              </p>
              <p style={{ margin: "5px 0" }}>
                <strong>Values:</strong>{" "}
                {response.parameters.values && response.parameters.values.length > 0 
                  ? response.parameters.values.map((v: any) => 
                      `${v.function}(${v.field})`
                    ).join(", ")
                  : "None (will use default count)"
                }
              </p>
            </div>
          )}
          
          {response.action === "chart" && response.parameters && (
            <div style={{ 
              backgroundColor: "#fff", 
              padding: "12px", 
              marginTop: "10px",
              borderRadius: "4px",
              border: "1px solid #e0e0e0"
            }}>
              <p style={{ margin: "5px 0" }}>
                <strong>Chart Type:</strong>{" "}
                <span style={{
                  backgroundColor: "#e8f5e9",
                  padding: "2px 8px",
                  borderRadius: "3px",
                  fontSize: "13px",
                  textTransform: "capitalize"
                }}>
                  {response.parameters.chartType}
                </span>
              </p>
              <p style={{ margin: "5px 0" }}>
                <strong>Title:</strong> {response.parameters.title}
              </p>
              <p style={{ margin: "5px 0" }}>
                <strong>Data Range:</strong>{" "}
                <code style={{
                  backgroundColor: "#f5f5f5",
                  padding: "2px 6px",
                  borderRadius: "3px"
                }}>
                  {response.parameters.dataRange}
                </code>
              </p>
              {response.parameters.xAxis && response.parameters.xAxis.title && (
                <p style={{ margin: "5px 0" }}>
                  <strong>X-Axis:</strong> {response.parameters.xAxis.title}
                </p>
              )}
              {response.parameters.yAxis && response.parameters.yAxis.title && (
                <p style={{ margin: "5px 0" }}>
                  <strong>Y-Axis:</strong> {response.parameters.yAxis.title}
                </p>
              )}
            </div>
          )}
          
          <button
            onClick={executeAction}
            style={{
              marginTop: "15px",
              padding: "12px 20px",
              backgroundColor: "#28a745",
              color: "white",
              border: "none",
              cursor: "pointer",
              width: "100%",
              fontSize: "15px",
              borderRadius: "4px",
              fontWeight: "600"
            }}
          >
            ✓ Execute in Excel
          </button>
        </div>
      )}
    </div>
  );
};

export default App;