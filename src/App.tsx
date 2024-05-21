import logo from './logo.svg';
import './App.css';
import React, { Component } from "react";
import { SpreadSheets } from "@mescius/spread-sheets-react";
import * as GC from "@mescius/spread-sheets";
import "@mescius/spread-sheets-charts";
import "@mescius/spread-sheets-shapes";
import "@mescius/spread-sheets-io";
import { saveAs } from "file-saver";

class App extends Component {
        hostStyle:any;
    spread: any;
    currentCustomerIndex: any;
    state: any;

  constructor(props:any) {
    super(props);
    this.hostStyle = {
      width: "1200px",
      height: "860px",
    };
    this.spread = null;
    this.state = {
      selectedFile: null,
    };
    this.currentCustomerIndex = 0;
  }
  render() {
    return (
      <div style={{ display: "flex" }}>
        <div style={{ width: "1200px", height: "860px" }}>
          <SpreadSheets
            workbookInitialized={(spread) => this.initSpread(spread)}
            hostStyle={this.hostStyle}
          ></SpreadSheets>
        </div>
        <div style={{ width: "200px", padding: "20px", background: "#ddd" }}>
          <p>Open Excel Files (.xlsx)</p>
          <input type="file" onChange={this.handleFileChange} />
          <button onClick={this.handleOpenExcel}>Open Excel</button>
          <p>Add Data</p>
          <button onClick={this.handleAddData}>Add Customer</button>
          <p>Save File</p>
          <button onClick={this.handleSaveExcel}>Save Excel</button>
        </div>
      </div>
    );
  }
  initSpread(spread: any) {
    this.spread = spread;
    this.spread.setSheetCount(2);
    console.log(this.spread);
  }
  handleSaveExcel = () => {
       var fileName = "Excel_Export.xlsx";
    this.spread.export(
      function (blob: any) {
        // save blob to a file
        saveAs(blob, fileName);
      },
      function (e: any) {
        console.log(e);
      },
      {
        fileType: GC.Spread.Sheets.FileType.excel,
      }
    );
  };
  handleOpenExcel = () => {
        const file = this.state.selectedFile;
    if (!file) {
      return;
    }
    // Specify the file type to ensure proper import
    const options = {
      fileType: GC.Spread.Sheets.FileType.excel,
    };
    this.spread.import(
      file,
      () => {
        console.log("Import successful");
      },
      (e: any) => {
        console.error("Error during import:", e);
      },
      options
    );
  };
  handleFileChange = (e: any) => {
    this.setState({
      selectedFile: e.target.files[0],
    });
  };
  handleAddData = () => {
     // Create new row and copy styles
    var newRowIndex = 34;
    var sheet = this.spread.getActiveSheet();
    sheet.addRows(newRowIndex, 1);
    sheet.copyTo(
      32,
      1,
      newRowIndex,
      1,
      1,
      11,
      GC.Spread.Sheets.CopyToOptions.style
    );
    // Define sample customer data
    var customerDataArrays = [
      ["Jessica Moth", 5000, 2000, 3000, 1300, 999, 100],
      ["John Doe", 6000, 2500, 3500, 1400, 1000, 20],
      ["Alice Smith", 7000, 3000, 4000, 1500, 1100, 0],
    ];
    // Get the current customer data array
    var currentCustomerData = customerDataArrays[this.currentCustomerIndex];
    // Add new data to the new row
    sheet.setArray(newRowIndex, 5, [currentCustomerData]);
    newRowIndex++;
    // Increment the index for the next button click
    this.currentCustomerIndex =
      (this.currentCustomerIndex + 1) % customerDataArrays.length;
  };
}

export default App;
