# TMLT to FHIR Converter

This Node.js application converts Thai Medical Laboratory Terminology (TMLT) data from an Excel file into a FHIR CodeSystem resource.

## Setup

1. Install dependencies:
   ```
   npm install
   ```

2. Place required input files in the `input` directory:
   - `TMLT-CS-template.json`: The template FHIR CodeSystem
   - `TMLTRF{version}.zip`: The zip file containing TMLT data (where {version} matches the version in config.js)

3. Configure the version in `config.js`:
   ```js
   module.exports = {
     version: '20250401' // YYYYMMDD format
   };
   ```

## Usage

Run the application:
```
npm start
```

This will:
1. Read the version from `config.js`
2. Extract data from the Excel file in the zip
3. Process each row and map it to the FHIR CodeSystem structure
4. Output the result to `output/CS-TMLT.js`
