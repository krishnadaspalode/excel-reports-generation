<h1 align="center">Excel Reports Generation</h1>

Create excel report from javascript which supports multiple sheets

A strigh forward way to create excel reports with multiple sheets in it.

## 🛠 Installation & Set Up

1. Clonse the repositiry
2. change to the directory 'excelreportsgeneration'
3. Load the 'index.html' file
4. Click on 'Generate Report' button


I have used "js-xlsx" bower component as dev dependency here

Library support 1: js-xlsx
```sh
bower install js-xlsx
```
Ref URL : https://libraries.io/bower/js-xlsx

I used the js-xlsx version-0.10.7 here in example and added below library too,

Library support 1: js-xlsx-style
```sh
bower install js-xlsx-style
```
Ref URL : https://libraries.io/bower/js-xlsx-style

We can use only the min.js files alone also in our application.
```sh
<script lang="javascript" src="xlsx.full.min.js"></script>
```
and 
```sh
<script lang="javascript"  src="jszip.js"></script>
```
and Finally the fileSaver.js file for export functionality,

<script src="FileSaver.js"></script>

The file can be obtained from the code or from below github

Ref URL : https://github.com/eligrey/FileSaver.js/tree/master/src


Steps to remmeber:

*   Specify the sheet structure
*   Specify column structre & width properties
*   Prepare data array
*   Create column headers
*   Iterate and insert data with relevant modifications as per requirement
*   Export report using fileSaver functionalities

## 📢 Result

<p align="center">
  <a href="https://github.com/krishnadaspalode/excelreportsgeneration" target="_blank">
    <img src="https://github.com/krishnadaspalode/excelreportsgeneration/result/view.png" alt="View" title="" width="60%">
    <br/>
    <br/>
    <img src="https://github.com/krishnadaspalode/excelreportsgeneration/result/result.png" alt="Report" title="" width="60%">
  </a>
</p>

## 💖 Support

If you are using this project and happy with it or just want to encourage me to continue creating stuff, you can do it by just starring and sharing the project.


## 💡 Contributing

Any contributors who want to make this project better can make contributions, which will be greatly appreciated. To contribute, clone this repo locally and commit your code to a new branch. Feel free to create an issue or make a pull request.