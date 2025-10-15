/**
 * displayDependencyGraph.js
 *
 * Tool: Display dependency graph in overlay modal
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "displayDependencyGraph",
  description: "Display dependency graph in an interactive overlay panel with navigation",
  executor: "frontend",
  parameters: {
    graphData: {
      type: "object",
      required: true,
      description: "Graph data from traceDependencyGraph tool (contains centerCell, nodes, edges, stats)"
    },
    displayOptions: {
      type: "object",
      required: false,
      description: "Display options",
      properties: {
        title: {
          type: "string",
          description: "Custom title for the panel"
        },
        highlightCritical: {
          type: "boolean",
          description: "Highlight cells with many connections"
        }
      }
    }
  },
  returns: {
    success: "boolean",
    displayed: "boolean"
  }
};

export async function execute(params) {
  console.log('🖼️ Displaying dependency graph in overlay:', params);

  const { graphData, displayOptions = {} } = params;

  // Validate graph data
  if (!graphData || !graphData.centerCell || !graphData.nodes) {
    throw new Error('Invalid graph data: missing required fields');
  }

  // Build and show modal
  showDependencyModal(graphData, displayOptions);

  return {
    success: true,
    displayed: true,
    panelType: "overlay"
  };
}

/**
 * Show dependency graph modal
 */
function showDependencyModal(graphData, options) {
  const modal = document.getElementById('dependencyModal');
  if (!modal) {
    console.error('Dependency modal not found in DOM');
    return;
  }

  // Build content
  const content = buildModalContent(graphData, options);
  
  // Insert content
  const modalBody = modal.querySelector('.dependency-modal-body');
  if (modalBody) {
    modalBody.innerHTML = content;
  }

  // Show modal
  modal.classList.remove('hidden');

  // Attach event handlers
  attachModalEventHandlers(graphData);
}

/**
 * Build modal content HTML
 */
function buildModalContent(graphData, options) {
  const { centerCell, nodes, edges, stats } = graphData;
  const title = options.title || `Dependencies for ${centerCell.address}`;

  let html = `
    <div class="dependency-header">
      <div class="dependency-title">
        <span class="dependency-icon">📊</span>
        <div>
          <h3>${title}</h3>
          <p class="dependency-subtitle">
            ${stats.totalNodes - 1} connected cells • 
            ${stats.precedentCount} precedents • 
            ${stats.dependentCount} dependents
          </p>
        </div>
      </div>
    </div>

    <div class="dependency-content">
      <!-- Center Cell -->
      <div class="dependency-section">
        <div class="section-label">Target Cell</div>
        <div class="trace-node center-node" 
             data-sheet="${centerCell.sheet}" 
             data-address="${centerCell.address}">
          <div class="node-icon">🎯</div>
          <div class="node-details">
            <div class="node-header">
              <span class="node-address">${centerCell.address}</span>
              <span class="node-sheet-badge">${centerCell.sheet}</span>
            </div>
            ${centerCell.formula ? `<div class="node-formula">${escapeHtml(centerCell.formula)}</div>` : ''}
            <div class="node-value">${escapeHtml(String(centerCell.value || ''))}</div>
          </div>
          <button class="node-nav-btn" data-action="goto">
            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M5 12h14M12 5l7 7-7 7" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
          </button>
        </div>
      </div>
  `;

  // Group nodes by type and level
  const precedents = nodes.filter(n => n.type === "precedent");
  const dependents = nodes.filter(n => n.type === "dependent");

  // Precedents section
  if (precedents.length > 0) {
    html += `
      <div class="dependency-section">
        <div class="section-label">
          <span>⬅ Precedents (${precedents.length})</span>
          <span class="section-hint">Cells that feed into ${centerCell.address}</span>
        </div>
        <div class="nodes-list">
    `;

    // Group by level
    const precedentsByLevel = groupByLevel(precedents);
    Object.keys(precedentsByLevel).sort((a, b) => Number(a) - Number(b)).forEach(level => {
      precedentsByLevel[level].forEach(node => {
        html += renderNodeCard(node, edges, centerCell, '⬅');
      });
    });

    html += `
        </div>
      </div>
    `;
  }

  // Dependents section
  if (dependents.length > 0) {
    html += `
      <div class="dependency-section">
        <div class="section-label">
          <span>➡ Dependents (${dependents.length})</span>
          <span class="section-hint">Cells that use ${centerCell.address}</span>
        </div>
        <div class="nodes-list">
    `;

    // Group by level
    const dependentsByLevel = groupByLevel(dependents);
    Object.keys(dependentsByLevel).sort((a, b) => Number(a) - Number(b)).forEach(level => {
      dependentsByLevel[level].forEach(node => {
        html += renderNodeCard(node, edges, centerCell, '➡');
      });
    });

    html += `
        </div>
      </div>
    `;
  }

  // No dependencies message
  if (precedents.length === 0 && dependents.length === 0) {
    html += `
      <div class="dependency-section">
        <div class="empty-state">
          <div class="empty-icon">🔍</div>
          <p>No dependencies found</p>
          <p class="empty-hint">This cell doesn't have any precedents or dependents</p>
        </div>
      </div>
    `;
  }

  html += `</div>`; // Close dependency-content

  return html;
}

/**
 * Render individual node card
 */
function renderNodeCard(node, edges, centerCell, arrow) {
  const isOtherSheet = node.sheet !== centerCell.sheet;
  const edge = edges.find(e => e.from === node.id || e.to === node.id);
  const levelBadge = node.level > 1 ? `<span class="level-badge">L${node.level}</span>` : '';

  return `
    <div class="trace-node" 
         data-sheet="${node.sheet}" 
         data-address="${node.address}">
      <div class="node-arrow">${arrow}</div>
      <div class="node-details">
        <div class="node-header">
          <span class="node-address">${node.address}</span>
          ${isOtherSheet ? `<span class="node-sheet-badge">${node.sheet}</span>` : ''}
          ${levelBadge}
        </div>
        ${node.formula ? `<div class="node-formula">${escapeHtml(node.formula)}</div>` : ''}
        <div class="node-value">${escapeHtml(String(node.value || ''))}</div>
      </div>
      <button class="node-nav-btn" data-action="goto" title="Navigate to this cell">
        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M5 12h14M12 5l7 7-7 7" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </button>
    </div>
  `;
}

/**
 * Attach event handlers to modal
 */
function attachModalEventHandlers(graphData) {
  const modal = document.getElementById('dependencyModal');

  // Close button
  const closeBtn = modal.querySelector('#dependencyModalClose');
  if (closeBtn) {
    closeBtn.onclick = () => {
      modal.classList.add('hidden');
    };
  }

  // Navigate buttons
  const navButtons = modal.querySelectorAll('[data-action="goto"]');
  navButtons.forEach(btn => {
    btn.onclick = async () => {
      const node = btn.closest('[data-sheet][data-address]');
      if (node) {
        const sheet = node.dataset.sheet;
        const address = node.dataset.address;
        await navigateToCell(sheet, address);
      }
    };
  });

  // Close on overlay click
  const overlay = modal.querySelector('.modal-overlay');
  if (overlay) {
    overlay.onclick = () => {
      modal.classList.add('hidden');
    };
  }
}

/**
 * Navigate to a specific cell in Excel
 */
async function navigateToCell(sheetName, address) {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(sheetName);
      const range = worksheet.getRange(address);
      
      // Activate worksheet
      worksheet.activate();
      
      // Select and scroll to range
      range.select();
      
      await context.sync();
      
      console.log(`✅ Navigated to ${sheetName}!${address}`);
    });
  } catch (error) {
    console.error('Failed to navigate:', error);
    alert(`Failed to navigate to ${sheetName}!${address}: ${error.message}`);
  }
}

/**
 * Group nodes by level
 */
function groupByLevel(nodes) {
  const grouped = {};
  nodes.forEach(node => {
    if (!grouped[node.level]) {
      grouped[node.level] = [];
    }
    grouped[node.level].push(node);
  });
  return grouped;
}

/**
 * Escape HTML to prevent XSS
 * Handles all types of inputs gracefully
 */
function escapeHtml(text) {
  try {
    // Handle null, undefined, or non-string values
    if (text === null || text === undefined) {
      return '';
    }
    
    // Convert to string if not already
    if (typeof text !== 'string') {
      text = String(text);
    }
    
    const map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    };
    
    return text.replace(/[&<>"']/g, m => map[m]);
  } catch (error) {
    // Fallback: return empty string if anything goes wrong
    console.warn('escapeHtml error:', error, 'Input:', text);
    return '';
  }
}

