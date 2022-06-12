import React, { Component } from 'react';

export class Home extends Component {
  static displayName = Home.name;

  render () {
    return (
      <div>
        <h1>Welcome to the ExcelGen application.</h1>
        <p>It is an application that uses next components:</p>
        <ul>
          <li><a href='https://get.asp.net/'>ASP.NET Core</a> and <a href='https://msdn.microsoft.com/en-us/library/67ef8sbd.aspx'>C#</a> for cross-platform server-side code</li>
          <li><a href='https://facebook.github.io/react/'>React</a> for client-side code</li>
          <li><a href='http://getbootstrap.com/'>Bootstrap</a> for layout and styling</li>
        </ul>
        <p>To use the application:</p>
        <ul>
          <li>Click at the 'Fetch data' button at top right corner.</li>
          <li>Click at the button 'Generate Excel'.</li>
          <li>Open downloaded excel file and see its filling.</li>
        </ul>
        <p style={{marginTop:'20%'}}>The application was built by <strong>Dmytro Moroziuk ⓒ Copyright 2022</strong></p>
      </div>
    );
  }
}
