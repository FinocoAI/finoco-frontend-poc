/**
 * traceDependencyGraph.js
 *
 * Tool: Recursively trace cell dependencies to build complete graph
 * Executor: Frontend (requires Office.js)
 */

import { parseFormulaReferences, getRepresentativeAddress } from '../../utils/formulaParser.js';

export const toolDefinition = {
  name: "traceDependencyGraph",
  description: "Trace cell dependencies recursively to build a complete dependency graph (precedents and/or dependents)",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet containing the cell"
    },
    address: {
      type: "string",
      required: true,
      description: "Cell address to trace (e.g., 'C10')"
    },
    direction: {
      type: "string",
      required: false,
      enum: ["precedents", "dependents", "both"],
      default: "both",
      description: "Direction to trace: 'precedents' (what feeds in), 'dependents' (what uses this), or 'both'"
    },
    maxDepth: {
      type: "number",
      required: false,
      default: 3,
      description: "Maximum depth to trace (1-5 recommended)"
    }
  },
  returns: {
    centerCell: "object (address, sheet, value, formula)",
    nodes: "array of node objects",
    edges: "array of edge objects",
    stats: "object with graph statistics"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing traceDependencyGraph:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, address, direction = "both", maxDepth = 3 } = params;

      // Validate params
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (!address || typeof address !== 'string') {
        throw new Error('Invalid address');
      }
      if (maxDepth < 1 || maxDepth > 5) {
        throw new Error('maxDepth must be between 1 and 5');
      }

      // Initialize graph structure
      const graph = {
        nodes: new Map(),
        edges: [],
        visited: new Set()
      };

      // Get center cell info
      const centerCell = await getCellInfo(context, sheetName, address);
      const centerNodeId = `${sheetName}!${address}`;
      
      // Add center node
      graph.nodes.set(centerNodeId, {
        id: centerNodeId,
        address: address,
        sheet: sheetName,
        value: centerCell.value,
        formula: centerCell.formula,
        level: 0,
        type: "center"
      });
      graph.visited.add(centerNodeId);

      console.log(`  📍 Center cell: ${centerNodeId}, value: ${centerCell.value}`);

      // Trace precedents (what feeds into this cell)
      if (direction === "precedents" || direction === "both") {
        console.log(`  ⬅ Tracing precedents...`);
        await tracePrecedentsRecursive(context, graph, sheetName, address, 1, maxDepth);
      }

      // Trace dependents (what uses this cell)
      if (direction === "dependents" || direction === "both") {
        console.log(`  ➡ Tracing dependents...`);
        await traceDependentsRecursive(context, graph, sheetName, address, 1, maxDepth);
      }

      // Convert nodes map to array
      const nodesArray = Array.from(graph.nodes.values());
      
      // Calculate stats
      const precedentNodes = nodesArray.filter(n => n.type === "precedent");
      const dependentNodes = nodesArray.filter(n => n.type === "dependent");

      const result = {
        centerCell: {
          address: address,
          sheet: sheetName,
          value: centerCell.value,
          formula: centerCell.formula
        },
        nodes: nodesArray,
        edges: graph.edges,
        stats: {
          totalNodes: nodesArray.length,
          precedentCount: precedentNodes.length,
          dependentCount: dependentNodes.length,
          edgeCount: graph.edges.length,
          maxDepth: maxDepth,
          direction: direction
        }
      };

      console.log(`  ✅ Graph complete: ${result.stats.totalNodes} nodes, ${result.stats.edgeCount} edges`);
      console.log(`     Precedents: ${result.stats.precedentCount}, Dependents: ${result.stats.dependentCount}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in traceDependencyGraph:`, error);
      throw {
        tool: "traceDependencyGraph",
        error: error.message,
        params: params
      };
    }
  });
}

/**
 * Recursively trace precedents (cells that feed into this cell)
 * Uses hybrid approach: Try Excel API first, fallback to formula parsing
 */
async function tracePrecedentsRecursive(context, graph, sheetName, address, level, maxDepth) {
  if (level > maxDepth) {
    return;
  }

  try {
    // Get worksheet and range
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const range = worksheet.getRange(address);
    range.load("address, formulas");
    await context.sync();

    const formula = range.formulas[0][0];
    
    console.log(`    ${'  '.repeat(level)}🔍 Checking ${sheetName}!${address}, formula: ${formula}`);
    
    // If no formula, no precedents
    if (!formula || !formula.startsWith('=')) {
      console.log(`    ${'  '.repeat(level)}❌ No formula found, skipping`);
      return;
    }

    let precedentRefs = [];
    let usedFormulaParsing = false;

    // APPROACH 1: Try Excel API first
    try {
      const precedents = range.getDirectPrecedents();
      precedents.load("address, areas");
      await context.sync();

      const areas = precedents.areas;
      areas.load("items");
      await context.sync();

      // Extract addresses from API
      for (let i = 0; i < areas.items.length; i++) {
        const area = areas.items[i];
        area.load("address");
        await context.sync();
        
        const fullAddress = area.address;
        const { sheet: precSheet, address: precAddr } = parseFullAddress(fullAddress, sheetName);
        precedentRefs.push({ sheet: precSheet, address: precAddr });
      }

      console.log(`    ${'  '.repeat(level)}📍 API found ${precedentRefs.length} precedents for ${sheetName}!${address}`);
    } catch (apiError) {
      // API failed - this is expected for complex formulas
      console.log(`    ${'  '.repeat(level)}⚠️ API failed for ${sheetName}!${address}: ${apiError.message}`);
      usedFormulaParsing = true;
    }

    // APPROACH 2: If API returned nothing or failed, parse formula
    if (precedentRefs.length === 0) {
      if (!usedFormulaParsing) {
        console.log(`    ${'  '.repeat(level)}📝 API returned 0 precedents, using formula parsing for ${sheetName}!${address}`);
      }
      
      console.log(`    ${'  '.repeat(level)}🔎 Parsing formula: ${formula}`);
      const parsedRefs = parseFormulaReferences(formula, sheetName);
      console.log(`    ${'  '.repeat(level)}📊 Parsed ${parsedRefs.length} references:`, parsedRefs);
      
      // Convert parsed references to addresses
      for (const ref of parsedRefs) {
        let addr = ref.address;
        
        // For full column/row references, use representative cell
        if (ref.isFullColumn || ref.isFullRow) {
          addr = getRepresentativeAddress(ref.address);
          console.log(`    ${'  '.repeat(level)}🔄 Converted ${ref.address} to ${addr} for visualization`);
        } else if (ref.isRange && !ref.isFullColumn && !ref.isFullRow) {
          // For regular ranges, use first cell
          addr = ref.address.split(':')[0];
        }
        
        precedentRefs.push({ sheet: ref.sheet, address: addr });
      }
      
      console.log(`    ${'  '.repeat(level)}✅ Final precedentRefs (${precedentRefs.length}):`, precedentRefs);
    }

    // Process all precedent references - with retry logic if API refs fail
    let retryWithParsing = false;
    let processedCount = 0;
    
    for (const { sheet: precSheet, address: precAddr } of precedentRefs) {
      const nodeId = `${precSheet}!${precAddr}`;

      // Skip if already visited
      if (graph.visited.has(nodeId)) {
        continue;
      }

      // Get cell info - with error handling
      let cellInfo;
      try {
        cellInfo = await getCellInfo(context, precSheet, precAddr);
      } catch (cellError) {
        console.warn(`    ${'  '.repeat(level)}⚠️ Could not get info for ${nodeId}: ${cellError.message}`);
        // If we got this from the API but can't access it, mark for retry
        if (!usedFormulaParsing) {
          retryWithParsing = true;
        }
        continue;
      }

      // Add node
      graph.nodes.set(nodeId, {
        id: nodeId,
        address: precAddr,
        sheet: precSheet,
        value: cellInfo.value,
        formula: cellInfo.formula,
        level: level,
        type: "precedent"
      });
      graph.visited.add(nodeId);

      // Add edge
      graph.edges.push({
        from: nodeId,
        to: `${sheetName}!${address}`,
        formula: cellInfo.formula || '',
        relationshipType: "feeds_into"
      });

      console.log(`    ${'  '.repeat(level)}⬅ Level ${level}: ${nodeId}`);
      processedCount++;

      // Recurse
      await tracePrecedentsRecursive(context, graph, precSheet, precAddr, level + 1, maxDepth);
    }
    
    // If API refs failed to load and we haven't used formula parsing yet, retry
    if (retryWithParsing && processedCount === 0) {
      console.log(`    ${'  '.repeat(level)}🔁 API references failed to load, retrying with formula parsing`);
      
      const parsedRefs = parseFormulaReferences(formula, sheetName);
      console.log(`    ${'  '.repeat(level)}📊 Parsed ${parsedRefs.length} references:`, parsedRefs);
      
      for (const ref of parsedRefs) {
        let addr = ref.address;
        
        if (ref.isFullColumn || ref.isFullRow) {
          addr = getRepresentativeAddress(ref.address);
          console.log(`    ${'  '.repeat(level)}🔄 Converted ${ref.address} to ${addr}`);
        } else if (ref.isRange && !ref.isFullColumn && !ref.isFullRow) {
          addr = ref.address.split(':')[0];
        }
        
        const nodeId = `${ref.sheet}!${addr}`;
        
        // Skip if already visited
        if (graph.visited.has(nodeId)) {
          continue;
        }
        
        // Get cell info
        try {
          const cellInfo = await getCellInfo(context, ref.sheet, addr);
          
          // Add node
          graph.nodes.set(nodeId, {
            id: nodeId,
            address: addr,
            sheet: ref.sheet,
            value: cellInfo.value,
            formula: cellInfo.formula,
            level: level,
            type: "precedent"
          });
          graph.visited.add(nodeId);

          // Add edge
          graph.edges.push({
            from: nodeId,
            to: `${sheetName}!${address}`,
            formula: cellInfo.formula || '',
            relationshipType: "feeds_into"
          });

          console.log(`    ${'  '.repeat(level)}⬅ Level ${level}: ${nodeId}`);

          // Recurse
          await tracePrecedentsRecursive(context, graph, ref.sheet, addr, level + 1, maxDepth);
        } catch (err) {
          console.warn(`    ${'  '.repeat(level)}⚠️ Could not process parsed ref ${nodeId}: ${err.message}`);
        }
      }
    }
  } catch (error) {
    // Silently handle errors for individual cells
    console.warn(`    Warning: Could not trace precedents for ${sheetName}!${address}:`, error.message);
  }
}

/**
 * Recursively trace dependents (cells that use this cell)
 */
async function traceDependentsRecursive(context, graph, sheetName, address, level, maxDepth) {
  if (level > maxDepth) {
    return;
  }

  try {
    // Get worksheet and range
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const range = worksheet.getRange(address);
    range.load("address");
    await context.sync();

    // Get direct dependents
    let dependents;
    try {
      dependents = range.getDirectDependents();
      dependents.load("address, areas");
      await context.sync();
    } catch (error) {
      // No dependents
      return;
    }

    // Extract dependent addresses
    const areas = dependents.areas;
    areas.load("items");
    await context.sync();

    for (let i = 0; i < areas.items.length; i++) {
      const area = areas.items[i];
      area.load("address");
      await context.sync();

      // Parse full address
      const fullAddress = area.address;
      const { sheet: depSheet, address: depAddr } = parseFullAddress(fullAddress, sheetName);
      const nodeId = `${depSheet}!${depAddr}`;

      // Skip if already visited
      if (graph.visited.has(nodeId)) {
        continue;
      }

      // Get cell info
      const cellInfo = await getCellInfo(context, depSheet, depAddr);

      // Add node
      graph.nodes.set(nodeId, {
        id: nodeId,
        address: depAddr,
        sheet: depSheet,
        value: cellInfo.value,
        formula: cellInfo.formula,
        level: level,
        type: "dependent"
      });
      graph.visited.add(nodeId);

      // Add edge
      graph.edges.push({
        from: `${sheetName}!${address}`,
        to: nodeId,
        formula: cellInfo.formula || '',
        relationshipType: "uses"
      });

      console.log(`    ${'  '.repeat(level)}➡ Level ${level}: ${nodeId}`);

      // Recurse
      await traceDependentsRecursive(context, graph, depSheet, depAddr, level + 1, maxDepth);
    }
  } catch (error) {
    // Silently handle errors for individual cells
    console.warn(`    Warning: Could not trace dependents for ${sheetName}!${address}:`, error.message);
  }
}

/**
 * Get cell information (value and formula)
 */
async function getCellInfo(context, sheetName, address) {
  const worksheet = context.workbook.worksheets.getItem(sheetName);
  const range = worksheet.getRange(address);
  
  range.load("values, formulas");
  await context.sync();

  return {
    value: range.values[0][0],
    formula: range.formulas[0][0]
  };
}

/**
 * Parse full cell address (handles both same-sheet and cross-sheet references)
 * Examples:
 *   "Sheet1!C10" -> { sheet: "Sheet1", address: "C10" }
 *   "C10" -> { sheet: defaultSheet, address: "C10" }
 */
function parseFullAddress(fullAddress, defaultSheet) {
  if (fullAddress.includes('!')) {
    const parts = fullAddress.split('!');
    return {
      sheet: parts[0].replace(/'/g, ''), // Remove quotes
      address: parts[1]
    };
  }
  return {
    sheet: defaultSheet,
    address: fullAddress
  };
}

