const fs = require('fs-extra');
const path = require('path');
const AdmZip = require('adm-zip');
const XLSX = require('xlsx');
const config = require('./config');

// Get the version from config
const version = config.version;

// Paths
const inputDir = path.join(__dirname, 'input');
const outputDir = path.join(__dirname, 'output');
const templatePath = path.join(inputDir, 'TMLT-CS-template.json');
const zipPath = path.join(inputDir, `TMLTRF${version}.zip`);
const outputPath = path.join(outputDir, 'CS-TMLT.js');

// Ensure output directory exists
fs.ensureDirSync(outputDir);

// Function to extract and process the XLSX file from zip
async function processZipFile() {
  try {
    console.log(`Processing zip file: ${zipPath}`);
    
    // Read the template file
    const templateContent = fs.readFileSync(templatePath, 'utf8');
    const template = JSON.parse(templateContent);
    
    // Extract the zip file
    const zip = new AdmZip(zipPath);
    const zipEntries = zip.getEntries();
    
    // Find the folder with the pattern TMLTRFYYYYMMDD
    const folderPattern = `TMLTRF${version}`;
    const tmltFolder = zipEntries.find(entry => 
      entry.entryName.startsWith(folderPattern) && entry.isDirectory);
    
    if (!tmltFolder) {
      throw new Error(`Folder matching pattern ${folderPattern} not found in zip file`);
    }
    
    console.log(`Found folder: ${tmltFolder.entryName}`);
    
    // Find the Excel file with the pattern TMLT_SNAPSHOTYYYYMMDD.xls
    const excelPattern = `TMLT_SNAPSHOT${version}.xls`;
    const excelEntry = zipEntries.find(entry => 
      entry.entryName.includes(excelPattern) && !entry.isDirectory);
    
    if (!excelEntry) {
      throw new Error(`Excel file matching pattern ${excelPattern} not found in zip file`);
    }
    
    console.log(`Found Excel file: ${excelEntry.entryName}`);
    
    // Extract the Excel file to a temporary location
    const tempDir = path.join(__dirname, 'temp');
    fs.ensureDirSync(tempDir);
    const tempExcelPath = path.join(tempDir, path.basename(excelEntry.entryName));
    
    zip.extractEntryTo(excelEntry, tempDir, false, true);
    
    // Read the Excel file
    const workbook = XLSX.readFile(tempExcelPath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    
    console.log(`Read ${data.length} rows from Excel file`);
    
    // Check for rows with empty properties
    let emptyComponentCount = 0;
    let emptyScaleCount = 0;
    let emptyUnitCount = 0;
    let emptySpecimenCount = 0;
    let emptyMethodCount = 0;
    
    data.forEach(row => {
      if (!row.COMPONENT) emptyComponentCount++;
      if (!row.SCALE) emptyScaleCount++;
      if (!row.UNIT) emptyUnitCount++;
      if (!row.SPECIMEN) emptySpecimenCount++;
      if (!row.METHOD) emptyMethodCount++;
    });
    
    console.log(`Rows with empty COMPONENT: ${emptyComponentCount}`);
    console.log(`Rows with empty SCALE: ${emptyScaleCount}`);
    console.log(`Rows with empty UNIT: ${emptyUnitCount}`);
    console.log(`Rows with empty SPECIMEN: ${emptySpecimenCount}`);
    console.log(`Rows with empty METHOD: ${emptyMethodCount}`);
    
    // Find a sample row with empty METHOD property
    const sampleWithEmptyMethod = data.find(row => !row.METHOD);
    
    if (sampleWithEmptyMethod) {
      console.log('Sample row with empty METHOD:');
      console.log(JSON.stringify(sampleWithEmptyMethod, null, 2));
    }
    
    // Remove the template concept (first one)
    template.concept = [];
    
    // Process each row and add to template
    data.forEach(row => {
      // Create properties array with non-empty values
      const properties = [];
      
      // TYPE is always included (even if empty)
      properties.push({
        code: "TYPE",
        valueCode: row.ORDER_TYPE || ""
      });
      
      // STATUS is always "active"
      properties.push({
        code: "STATUS",
        valueCode: "active"
      });
      
      // Only add properties that have values
      if (row.COMPONENT) {
        properties.push({
          code: "COMPONENT",
          valueString: row.COMPONENT
        });
      }
      
      if (row.SCALE) {
        properties.push({
          code: "SCALE",
          valueString: row.SCALE
        });
      }
      
      if (row.UNIT) {
        properties.push({
          code: "UNIT",
          valueString: row.UNIT
        });
      }
      
      if (row.SPECIMEN) {
        properties.push({
          code: "SPECIMEN",
          valueString: row.SPECIMEN
        });
      }
      
      if (row.METHOD) {
        properties.push({
          code: "METHOD",
          valueString: row.METHOD
        });
      }
      
      const concept = {
        code: row.TMLT_Code,
        display: row.TMLT_Name,
        property: properties
      };
      
      template.concept.push(concept);
    });
    
    // Update template version with the one from config
    template.version = version;
    template.title = `Thai Medical Laboratory Terminology (TMLT) ${version}`;
    template.date = `${version.substring(0, 4)}-${version.substring(4, 6)}-${version.substring(6, 8)}T00:00:00+07:00`;
    
    // Write the result to a JS file
    const jsContent = `// Generated TMLT CodeSystem - Version ${version}
module.exports = ${JSON.stringify(template, null, 2)};`;
    
    fs.writeFileSync(outputPath, jsContent, 'utf8');
    console.log(`Output written to ${outputPath}`);
    
    // Clean up
    fs.removeSync(tempDir);
    
  } catch (error) {
    console.error('Error processing the zip file:', error);
  }
}

// Run the process
processZipFile(); 