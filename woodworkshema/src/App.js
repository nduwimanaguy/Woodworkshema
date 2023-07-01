import React from 'react';
import * as XLSX from 'xlsx';
import FileSaver from 'file-saver';
import logo from './logo.svg';
import './App.css';

function App() {
  const handleDownloadExcel = () => {
    /* Read the existing workbook */
    const workbook = XLSX.readFile('xfactuurtemplate.xlsx');
    /* Get the first worksheet */
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    /* Modify the worksheet as needed */
    worksheet.A1.v = 'Hello, world!';
    /* Write the modified workbook back to the file */
    XLSX.writeFile(workbook, '/public/xfactuurtemplate.xlsx');
  };

  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          Edit <code>src/App.js</code> and save to reload.
        </p>
        <a
          className="App-link"
          href="https://reactjs.org"
          target="_blank"
          rel="noopener noreferrer"
        >
          Learn React
        </a>
        <button onClick={handleDownloadExcel}>Download Excel</button>
      </header>
    </div>
  );
}

export default App;
