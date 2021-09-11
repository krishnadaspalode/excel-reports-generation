<h1 align="center">Excel Reports Generation</h1>

Create excel report from javascript which supports multiple sheets

<p align="center">
  <a href="https://github.com/krishnadaspalode/excelreportsgeneration" target="_blank">
    <img src="https://github.com/krishnadaspalode/excelreportsgeneration/tree/main/result/view.png" alt="Preview" title="" width="60%">
    <br/>
    <br/>
    <img src="https://github.com/krishnadaspalode/excelreportsgeneration/tree/main/result.png" alt="Preview" title="" width="60%">
  </a>
</p>

A strigh forward way to create excel reports with multiple sheets in it.

I have used "js-xlsx" bower component as dev dependency here and FileSaver.js for exporting file.

Library support 1: js-xlsx

bower install js-xlsx

Ref URL : https://libraries.io/bower/js-xlsx

I used the js-xlsx version-0.10.7 here in example and add below library too,

Library support 1: js-xlsx-style

bower install js-xlsx-style

Ref URL : https://libraries.io/bower/js-xlsx-style

We can use only the min.js files alone also in our application.

<script lang="javascript" src="xlsx.full.min.js"></script>
&
<script lang="javascript"  src="jszip.js"></script>

and Finally the fileSaver.js file for export functionality,

<script src="FileSaver.js"></script>

The script file can be obtained from the code or from below github

Ref URL : https://github.com/eligrey/FileSaver.js/tree/master/src


Steps to remmeber:

*   Specify the sheet structure
*   Specify column structre & width properties
*   Prepare data array
*   Create column headers
*   Iterate and insert data with relevant modifications as per requirement
*   Export report using fileSaver functionalities



