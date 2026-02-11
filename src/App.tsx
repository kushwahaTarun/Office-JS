import "./App.css";
import ExcelComponent from "./components/ExcelComponent";

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h1>Fluid AI GPT - Excel Add-in</h1>
        <p>Office.js Excel Integration</p>
      </header>
      <main>
        <ExcelComponent />
      </main>
    </div>
  );
}

export default App;
