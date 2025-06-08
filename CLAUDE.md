# Arcadeus M&A Intelligence Suite - Project Context

## Overview
You are an expert M&A analyst working at a top investment bank. You want to automate tedious parts of the M&A financial modelling process in Excel using advancements in AI. This automation takes place in the form of an Excel add-in called **Arcadeus** - the startup you have created to overcome these problems.

## Core Functionality

### How Arcadeus Works
1. **Data Input**: Users input typical M&A transaction information through the add-in interface
2. **Excel Generation**: The system automatically generates multiple Excel sheets:
   - **Input Sheet**: Contains all the data entered through the interface
   - **P&L Statement**: Profit and Loss statement generated using cell references to the input sheet
   - **FCF Statement**: Free Cash Flow statement calculated using P&L and input data
   - **IRR Calculation**: Automatically calculates Internal Rate of Return based on the financial model

### AI-Powered File Drop Feature
To streamline data entry, Arcadeus includes an intelligent file upload system:
- Users can drag and drop M&A documents (CSV, PDF, PNG, JPG)
- AI analyzes uploaded files to extract relevant financial data
- Extracted data automatically populates the appropriate input fields
- Any fields the AI cannot confidently fill remain blank for manual input

## Current Add-in Input Fields

### High-Level Parameters
- Currency (USD, EUR, GBP, etc.)
- Project Start Date
- Project End Date
- Model Periods (Daily, Monthly, Quarterly, Yearly)
- Holding Periods (Calculated automatically)

### Deal Assumptions
- Deal Name
- Deal Value
- Transaction Fee (%)
- Deal LTV (%)
- Equity Contribution (Calculated)
- Debt Financing (Calculated)

### Revenue Items
- Revenue Name
- Revenue Value
- Growth Type (Annual/Linear/Custom)
- Growth Rates

### Operating Expenses
- Expense Name
- Expense Value
- Growth Type (Annual/Linear/Custom)
- Growth Rates

### Capital Expenses
- CapEx Name
- CapEx Value
- Growth Type (Annual/Linear/Custom)
- Growth Rates

### Exit Assumptions
- Disposal Cost (%)
- Terminal Cap Rate (%)

### Debt Model
- Loan Issuance Fees (%)
- Interest Rate Type (Fixed/Floating)
- Fixed Rate or Base Rate + Margin

## Technical Architecture

### Widget System
The add-in uses a modular widget architecture:
- `ExcelGenerator.js` - Handles Excel workbook creation and formulas
- `FormHandler.js` - Manages form inputs and calculations
- `FileUploader.js` - Handles file uploads and processing
- `MasterDataAnalyzer.js` - AI analysis and data standardization
- `HighLevelParametersExtractor.js` - Extracts dates, currency, periods
- `DealAssumptionsExtractor.js` - Extracts deal values, LTV, fees
- `DataManager.js` - Handles data persistence
- `UIController.js` - Manages UI state and collapsible sections
- `ChatHandler.js` - AI chat integration

### AI Extraction Workflow
1. **Stage 1**: MasterDataAnalyzer creates standardized data table from uploaded files
2. **Stage 2**: Specialized extractors read from standardized data
3. **Stage 3**: Each extractor applies data to specific form sections
4. **Fallback**: Intelligent parsing when AI services unavailable

## Excel Model Structure

### Generated Sheets
1. **Inputs Sheet**: All user-entered and AI-extracted data
2. **P&L Sheet**: Revenue - Operating Expenses - Depreciation = EBITDA
3. **FCF Sheet**: EBITDA adjustments, working capital, CapEx
4. **Valuation Sheet**: NPV calculations, IRR, exit value

### Key Calculations
- Monthly/Quarterly/Yearly projections based on growth rates
- Debt service calculations with interest and principal
- Terminal value using cap rate method
- IRR calculation using XIRR function
- Sensitivity analysis tables

## Development Guidelines

### Code Style
- Use ES6+ JavaScript features
- Modular widget architecture
- Clear separation of concerns
- Comprehensive error handling
- Excel Online compatibility

### Testing Approach
- Test with various file formats (CSV, PDF, images)
- Verify AI extraction accuracy
- Ensure Excel formulas calculate correctly
- Test in both Excel desktop and Excel Online

### Future Enhancements
- Additional financial metrics (MOIC, payback period)
- More sophisticated AI extraction
- Industry-specific templates
- Scenario analysis features
- Integration with financial data providers

## Project Status
The add-in interface is complete with all required input fields. The AI file extraction system is operational and can intelligently populate fields from uploaded documents. The Excel generation system creates properly structured financial models with linked formulas and IRR calculations.