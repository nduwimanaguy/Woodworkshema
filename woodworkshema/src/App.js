import React from 'react';
import * as XLSX from 'xlsx';
import FileSaver from 'file-saver';

function App() {
  const handleEditExcel = () => {
    // Read the Excel file from the public/ directory
    const excelFilePath = process.env.PUBLIC_URL + '/xfactuurtemplate.xlsx';

    fetch(excelFilePath)
      .then((response) => response.arrayBuffer())
      .then((data) => {
        /* Read the existing workbook */
        const wb = XLSX.read(data, { type: 'array' });

        /* Get the first worksheet from the workbook */
        const ws = wb.Sheets[wb.SheetNames[0]];

        /* Edit the value in cell A1 */
        XLSX.utils.sheet_add_aoa(ws, [['Hello, world! (edited)']], { origin: 'A1' });

        /* Write the updated workbook to a file */
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        FileSaver.saveAs(blob, 'xfactuurtemplate_edited.xlsx');
      })
      .catch((error) => {
        console.error('Error reading the file:', error);
      });
  };

  return (
    <div className="App">
      <header className="App-header">
        {/* Your app UI */}
        <button onClick={handleEditExcel}>Edit Excel</button>
      </header>
    </div>
  );
}

export default App;
