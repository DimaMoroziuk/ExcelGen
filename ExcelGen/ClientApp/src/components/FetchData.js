import React, { Component } from 'react';

export class FetchData extends Component {

  constructor(props) {
    super(props);
    this.state = {isHovering: false}
  }

  // componentDidMount() {
  //   this.SaveExcel()
  // }

  handleMouseEnter = () => {
    this.state.isHovering = true;
  };

  handleMouseLeave = () => {
    this.state.isHovering = false;
  };

  render() {

    return (
      <div>
        {/* <a id="downloadExcelLink" download="excelFile.xlsx" href="blob:http://localhost:3000/api/Discipline/excelTry">Download Excel File</a> */}
        <h1 style={{marginTop: '5%'}}  id="tabelLabel" >Excel Generation</h1>
        <p>The page will give you opportunity to create excel file with disciplines. Try it!!!</p>

        <button 
        onClick={() => this.populateExcel()}
        style={{marginTop: '35%',
                borderColor: 'black',
                backgroundColor: this.state.isHovering ? 'black' : 'white',
                color: this.state.isHovering ? 'white' : 'black',
                borderRadius: '5px'}} 
                onMouseEnter={this.handleMouseEnter}
                onMouseLeave={this.handleMouseLeave}
        className='btn'>
        Generate  Excel
    </button>
    {/* <a href = 'C:\ExcelDemo.xlsx' download="foo.xlsx" target="_blank">Or try to use the link...</a> */}
      </div>
    );
  }

  async populateExcel() {
    const downloadExcelResponse = await fetch('api/Discipline/excelGen');
  }

  // async SaveExcel() {
  //   const downloadExcelResponse = await fetch('api/Discipline/excelGenSaved');
  // }
}
