import { useEffect, useState } from "react";

export default function ExcelComponent() {
  const [value, setValue] = useState<number | string>("");
  const [inputValue, setInputValue] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);
  const [address, setAddress] = useState<string>("");

  useEffect(() => {
    loadSelectedCell();
  }, []);

  const loadSelectedCell = async () => {
    try {
      setLoading(true);
      setError("");

      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "address"]);
        await context.sync();

        // Get the selected cell value
        const cellValue = range.values[0][0];
        setValue(cellValue ?? "");
        setAddress(range.address);
        setInputValue(String(cellValue ?? ""));
      });
    } catch (err) {
      const errorMessage =
        err instanceof Error ? err.message : "Failed to read cell value";
      setError(errorMessage);
      console.error("Error loading cell:", err);
    } finally {
      setLoading(false);
    }
  };

  const writeToCell = async () => {
    try {
      setLoading(true);
      setError("");

      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.values = [[inputValue]];
        await context.sync();

        setValue(inputValue);
      });
    } catch (err) {
      const errorMessage =
        err instanceof Error ? err.message : "Failed to write cell value";
      setError(errorMessage);
      console.error("Error writing to cell:", err);
    } finally {
      setLoading(false);
    }
  };

  const handleRefresh = () => {
    loadSelectedCell();
  };

  return (
    <div style={{ padding: "20px", maxWidth: "500px", margin: "0 auto" }}>
      <h2>Excel Cell Manager</h2>

      {error && (
        <div
          style={{
            padding: "10px",
            backgroundColor: "#fee",
            color: "#c00",
            borderRadius: "4px",
            marginBottom: "15px",
          }}
        >
          <strong>Error:</strong> {error}
        </div>
      )}

      <div style={{ marginBottom: "20px" }}>
        <h3>Current Selection</h3>
        {loading ? (
          <p>Loading...</p>
        ) : (
          <>
            <p>
              <strong>Address:</strong> {address || "No cell selected"}
            </p>
            <p>
              <strong>Value:</strong> {String(value)}
            </p>
          </>
        )}
        <button
          onClick={handleRefresh}
          disabled={loading}
          style={{
            padding: "8px 16px",
            backgroundColor: "#0078d4",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading ? "not-allowed" : "pointer",
            opacity: loading ? 0.6 : 1,
          }}
        >
          Refresh Selection
        </button>
      </div>

      <div style={{ marginBottom: "20px" }}>
        <h3>Write to Cell</h3>
        <input
          type="text"
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          placeholder="Enter value"
          disabled={loading}
          style={{
            width: "100%",
            padding: "8px",
            marginBottom: "10px",
            border: "1px solid #ccc",
            borderRadius: "4px",
            fontSize: "14px",
          }}
        />
        <button
          onClick={writeToCell}
          disabled={loading}
          style={{
            padding: "8px 16px",
            backgroundColor: "#107c10",
            color: "white",
            border: "none",
            borderRadius: "4px",
            cursor: loading ? "not-allowed" : "pointer",
            opacity: loading ? 0.6 : 1,
          }}
        >
          {loading ? "Writing..." : "Write to Selected Cell"}
        </button>
      </div>

      <div
        style={{
          padding: "10px",
          backgroundColor: "#f0f0f0",
          borderRadius: "4px",
          fontSize: "12px",
        }}
      >
        <p>
          <strong>Instructions:</strong>
        </p>
        <ol style={{ marginLeft: "20px" }}>
          <li>Select a cell in Excel</li>
          <li>View the current cell value</li>
          <li>Enter a new value and click "Write to Selected Cell"</li>
          <li>Use "Refresh Selection" to update the displayed value</li>
        </ol>
      </div>
    </div>
  );
}
