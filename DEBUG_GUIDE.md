# DEBUG GUIDE - Supplier Mapping

## How to Use the Debug Feature

The debug system helps you track why specific suppliers are not calculating correctly by showing the complete mapping process from COD IMS/COD FORNECEDOR to BD_PN.

### Quick Start

1. **Open DB.py** and locate the debug configuration at the top (around line 70-80)

2. **Add your supplier codes** to the `DEBUG_SUPPLIERS` array:

```python
# ==================== DEBUG CONFIGURATION ====================
# Add supplier codes here to debug mapping issues
DEBUG_SUPPLIERS = ['33611', '800030798']  # Add your supplier codes here

# Optionally filter by AGRUPAMENTO (leave empty to see all)
DEBUG_AGRUPAMENTO = ['30956207']  # Add AGRUPAMENTO codes to filter debug output
```

3. **Run your program** normally

4. **Check the console output** for detailed debug information

## ⚡ CRITICAL FIX: Consistent Code Usage

**Problem Solved**: Previously, the system might use COD FORNECEDOR for some lookups and COD IMS for others, causing inconsistent results even when data existed in BD_PN. For example, QME might be found using COD IMS, but FORNECEDOR mapping failed using COD FORNECEDOR.

**New Solution**: The system now:
1. **Tests BOTH codes** (COD FORNECEDOR and COD IMS) against BD_PN at the start
2. **Determines which code actually works** for each supplier
3. **Uses ONLY that code consistently** for ALL subsequent lookups (FORNECEDOR, DESCRIÇÃO, QME, PESO, etc.)

This ensures if a supplier is registered under COD IMS in BD_PN, **all lookups** use COD IMS. If registered under COD FORNECEDOR, **all lookups** use COD FORNECEDOR.

Example debug output showing code determination:
```
[SUPPLIER_CODE_DETERMINATION] Determined supplier code for DESENHO 522343060
  - AGRUPAMENTO: 30956207
  - COD FORNECEDOR: 800030798
  - COD IMS: 33611
  - Code that works: 33611
  - Source: COD IMS
  - Status: ✓ Found in BD_PN - will use COD IMS for all lookups

[QME_RESOLUTION] Resolving QME for DESENHO 522343060
  - Working Code: 33611
  - Code Source: COD_IMS
  - Resolution attempt: Code=33611, Key=33611|522343060_E197, Found=True, Value=6
  - Final result: 6
  - Note: Using COD_IMS consistently (determined from BD_PN lookup)
```

Notice how once COD IMS (33611) is determined to work, it's used for ALL lookups!

### Example: Debugging Supplier with Multiple Codes

**Important**: A single supplier can have BOTH COD IMS and COD FORNECEDOR!

For example, the same supplier might have:
- COD IMS = 33611
- COD FORNECEDOR = 800030798

To debug this supplier, add EITHER code (or both):
```python
DEBUG_SUPPLIERS = ['33611']           # Track by COD IMS
# OR
DEBUG_SUPPLIERS = ['800030798']       # Track by COD FORNECEDOR  
# OR
DEBUG_SUPPLIERS = ['33611', '800030798']  # Track by either code
```

The system will detect when either code appears and show you the complete mapping process.

### What Gets Tracked

The debug system tracks these stages for each supplier:

**0. SUPPLIER_CODE_DETERMINATION** (New!)
   - **Critical**: Determines which code (COD FORNECEDOR or COD IMS) works for this supplier
   - Tests both codes against BD_PN to see which one exists
   - **This code is then used consistently for ALL subsequent lookups**
   - Shows which code was found and from which source

1. **MAP_KEY_GENERATION** (Deprecated - replaced by SUPPLIER_CODE_DETERMINATION)
   - Shows the original COD IMS and COD FORNECEDOR values
   - Shows the MAP_KEY that was generated
   - Explains the logic used (COD FORNECEDOR first, then COD IMS, taking first code before "/")

2. **FORNECEDOR_MAPPING**
   - Shows if the FORNECEDOR name was successfully mapped from BD_PN
   - Shows the MAP_KEY used for lookup
   - Indicates if mapping failed (MAP_KEY not found in BD_PN)

3. **DESCRIÇÃO_MATERIAL**
   - Shows all candidate codes tried (COD FORNECEDOR first, then COD IMS)
   - Shows each lookup attempt with the full key used
   - Shows whether each attempt found a match
   - Shows the final result

4. **DESCRIÇÃO_EMBALAGEM**
   - Similar to DESCRIÇÃO_MATERIAL but for MDR descriptions
   - Uses COD FORNECEDOR + MDR for lookups, then tries COD IMS

5. **QME_RESOLUTION**
   - Shows how QME (quantity per package) is resolved
   - Critical for calculations - if QME is not found, calculations fail
   - Shows all lookup attempts and the final result
   - Tries COD FORNECEDOR first, then COD IMS

### Debug Output Format

```
[DEBUG] Supplier: 800030798 | Stage: MAP_KEY_GENERATION
        Generating MAP_KEY for DESENHO 12345
        COD IMS (raw): 33611
        COD FORNECEDOR (raw): 800030798
        MAP_KEY (generated): 800030798
        Logic: Using COD FORNECEDOR first, else COD IMS, taking first code before "/"

[DEBUG] Supplier: 800030798 | Stage: FORNECEDOR_MAPPING
        Mapping FORNECEDOR for DESENHO 12345 - SUCCESS
        MAP_KEY (numeric): 800030798.0
        FORNECEDOR (mapped): SUPPLIER NAME
        Status: ✓ Mapped successfully
```

### At the End - Debug Summary

After processing completes, a comprehensive summary is printed showing all debug information grouped by supplier. This makes it easy to review the entire mapping process.

### Common Issues and Solutions

#### Issue: FORNECEDOR_MAPPING Failed
**Cause**: The MAP_KEY (derived from COD FORNECEDOR or COD IMS) doesn't exist in BD_PN's COD FORNECEDOR column

**Solution**: 
- Check if the supplier code exists in BD_PN
- The system tries COD FORNECEDOR first, then COD IMS
- Verify which code is actually registered in BD_PN for this supplier

#### Issue: DESCRIÇÃO_MATERIAL Not Found
**Cause**: None of the attempted keys (COD FORNECEDOR|KEY or COD IMS|KEY) exist in BD_PN

**Solution**:
- Check which code (COD FORNECEDOR or COD IMS) is used for this supplier in BD_PN
- Verify the DESENHO and MDR combination exists in BD_PN
- Check if there are multiple codes separated by "/" - the system tries each one
- Remember: same supplier can have both COD FORNECEDOR and COD IMS - system tries both

#### Issue: QME Not Found (Critical!)
**Cause**: QME is missing from BD_PN for this supplier/DESENHO/MDR combination

**Solution**:
- QME MUST be present for calculations to work
- Add QME to BD_PN for this combination
- Without QME, QTD EMBALAGENS will be 0 and calculations will fail

### Key Insight

**In BD_PN, some suppliers use COD IMS and others use COD FORNECEDOR!**

Additionally, **a single supplier can have BOTH codes** (e.g., COD IMS = 33611, COD FORNECEDOR = 800030798).

**NEW BEHAVIOR** (Consistent Code Usage):
1. System tests BOTH COD FORNECEDOR and COD IMS to see which one exists in BD_PN
2. Once determined, uses ONLY that code for ALL lookups
3. No more mixing codes - consistency guaranteed!

**Why this matters**: 
- If a supplier's QME is registered under COD IMS but you try to look it up with COD FORNECEDOR, it will fail
- The new system ensures the SAME code is used everywhere, eliminating these mismatches
- Items that previously showed QME but didn't calculate will now calculate correctly

Use the debug output to see which code was determined to work for your specific supplier, and verify all lookups use that same code.

### Tips

- Start with a small set of suppliers to debug (1-3 at a time)
- Review the debug summary at the end for a complete picture
- Pay special attention to QME - it's critical for calculations
- Compare successful mappings with failed ones to identify patterns
