import React from 'react';
import * as XLSX from 'xlsx';
import FileSaver from 'file-saver';
import logo from './logo.svg';
import './App.css';

function App() {
  const handleDownloadExcel = () => {
    /* Create a new workbook */
    const wb = XLSX.utils.book_new();
    /* Create a new worksheet */
    const ws = XLSX.utils.aoa_to_sheet([['Hello, world!']]);
    /* Add the worksheet to the workbook */
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    /* Write the workbook to a file */
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    FileSaver.saveAs(blob, 'hello_world.xlsx');
  };

  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  }

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
