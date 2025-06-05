# Cost Items Extraction Test Examples

## Test CSV Files for Cost Extraction

### Example 1: Simple Cost Items
```csv
Cost Item 1,200000
Cost Item 2,1000
Cost Item 3,20000
OpEx Cost Inflation,2
CapEx Cost Inflation,1.5
```
**Expected Output:**
- Cost Item 1: $200,000 (linear 2%)
- Cost Item 2: $1,000 (linear 2%)  
- Cost Item 3: $20,000 (linear 1.5%)

### Example 2: Staff Expenses with Growth
```csv
Staff expenses,60000
Salary Growth (p.a.),0.5
Marketing costs,25000
Office rent,48000
General inflation,3
```
**Expected Output:**
- Staff expenses: $60,000 (linear 0.5%)
- Marketing costs: $25,000 (linear 3%)
- Office rent: $48,000 (linear 3%)

### Example 3: Mixed OpEx and CapEx
```csv
Staff expenses,60000
Salary Growth (p.a.),0.50
Cost Item 1,200000
Cost Item 2,1000
OpEx Cost Inflation,2.00
Cost Item 3,20000
Cost Item 4,3000
CapEx Cost Inflation,1.50
```
**Expected Output:**
- Staff expenses: $60,000 (linear 0.5%)
- Cost Item 1: $200,000 (OpEx - linear 2%)
- Cost Item 2: $1,000 (OpEx - linear 2%)
- Cost Item 3: $20,000 (CapEx - linear 1.5%)
- Cost Item 4: $3,000 (CapEx - linear 1.5%)

### Example 4: No Growth Specified
```csv
Utilities,12000
Insurance,8000
Legal fees,15000
```
**Expected Output:**
- Utilities: $12,000 (no_growth)
- Insurance: $8,000 (no_growth)
- Legal fees: $15,000 (no_growth)

## Testing Instructions

1. Create a CSV with any of the above examples
2. Upload to the Excel add-in
3. Click "Auto Fill with AI"
4. Check the success message for "Cost Items:" section
5. Verify the Cost Items section in your form shows the extracted items

## Expected Success Message Format

After extraction, you should see:
```
✅ Data Extraction Successful!

Cost Items:
• Staff expenses: $60,000 (linear 0.5%)
• Cost Item 1: $200,000 (linear 2%)
• Cost Item 2: $1,000 (no_growth)
```