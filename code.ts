// Show the plugin UI
figma.showUI(__html__, { width: 450, height: 550 });

// Message handler types
type PopulateTableMessage = {
  type: 'populate-table';
  data: any;
  format: 'objects';
};

type GenerateDummyDataMessage = {
  type: 'generate-dummy-data';
  rowCount: number;
  columnCount: number;
};

type CancelMessage = {
  type: 'cancel';
};

type PluginMessage = PopulateTableMessage | GenerateDummyDataMessage | CancelMessage;

// Listen for messages from the UI
figma.ui.onmessage = async (msg: PluginMessage) => {
  if (msg.type === 'populate-table') {
    await populateTable(msg.data);
  } else if (msg.type === 'generate-dummy-data') {
    await populateTableWithDummyData(msg.rowCount, msg.columnCount);
  } else if (msg.type === 'cancel') {
    figma.closePlugin();
  }
};

// Load fonts before editing text
async function loadFonts(nodes: SceneNode[]): Promise<void> {
  const textNodes = nodes.filter(node => node.type === 'TEXT') as TextNode[];
  const fontNames = new Set<string>();
  
  textNodes.forEach(node => {
    if (node.fontName !== figma.mixed) {
      fontNames.add(JSON.stringify(node.fontName));
    }
  });
  
  const fontLoadPromises = Array.from(fontNames).map(fontStr => {
    const font = JSON.parse(fontStr) as FontName;
    return figma.loadFontAsync(font);
  });
  
  await Promise.all(fontLoadPromises);
}

// Find value layers within a cell component
function findValueLayer(node: SceneNode): TextNode | null {
  if (node.name.toLowerCase() === 'value' && node.type === 'TEXT') {
    return node as TextNode;
  }
  
  if ('children' in node) {
    const parent = node as ChildrenMixin & SceneNode;
    for (const child of parent.children) {
      const value = findValueLayer(child);
      if (value) return value;
    }
  }
  
  return null;
}

// Get all cell instances from a row or direct selection
function getAllCells(selection: readonly SceneNode[]): InstanceNode[] {
  const cells: InstanceNode[] = [];
  
  selection.forEach(node => {
    // If the node is a component instance that might be a cell
    if (node.type === 'INSTANCE') {
      if (node.name.toLowerCase() === 'cell') {
        cells.push(node);
      } else {
        // This might be a row, so check its children
        if ('children' in node) {
          const parent = node as ChildrenMixin & SceneNode;
          const cellInstances = parent.children.filter(child => 
            child.type === 'INSTANCE' && child.name.toLowerCase() === 'cell'
          ) as InstanceNode[];
          
          cells.push(...cellInstances);
        }
      }
    } 
    // If the node might be a row or container with cells
    else if ('children' in node) {
      const parent = node as ChildrenMixin & SceneNode;
      
      // First level - might be rows
      parent.children.forEach(child => {
        if (child.type === 'INSTANCE') {
          if (child.name.toLowerCase() === 'cell') {
            cells.push(child);
          } else if ('children' in child) {
            // This might be a row, check its children
            const rowNode = child as InstanceNode & ChildrenMixin;
            const cellInstances = rowNode.children.filter(grandchild => 
              grandchild.type === 'INSTANCE' && grandchild.name.toLowerCase() === 'cell'
            ) as InstanceNode[];
            
            cells.push(...cellInstances);
          }
        } else if ('children' in child) {
          // This might be a container with cells
          const containerNode = child as SceneNode & ChildrenMixin;
          const cellInstances = containerNode.children.filter(grandchild => 
            grandchild.type === 'INSTANCE' && grandchild.name.toLowerCase() === 'cell'
          ) as InstanceNode[];
          
          cells.push(...cellInstances);
        }
      });
    }
  });
  
  return cells;
}

// Organize cells into rows
function organizeCellsIntoRows(cells: InstanceNode[]): InstanceNode[][] {
  // Group cells by their parent (row)
  const rowMap = new Map<string, InstanceNode[]>();
  
  cells.forEach(cell => {
    if (cell.parent) {
      const parentId = cell.parent.id;
      if (!rowMap.has(parentId)) {
        rowMap.set(parentId, []);
      }
      rowMap.get(parentId)?.push(cell);
    }
  });
  
  // Sort cells within each row by their x position
  rowMap.forEach(rowCells => {
    rowCells.sort((a, b) => a.x - b.x);
  });
  
  // Convert map to array of rows
  const rows: InstanceNode[][] = Array.from(rowMap.values());
  
  // Sort rows by their y position (top to bottom)
  rows.sort((rowA, rowB) => {
    if (rowA.length === 0 || rowB.length === 0) return 0;
    return rowA[0].y - rowB[0].y;
  });
  
  return rows;
}

// Convert array of objects to a 2D array format
function convertObjectsTo2DArray(data: any[]): any[][] {
  if (!Array.isArray(data) || data.length === 0) {
    return [];
  }
  
  // Extract keys from the first object to use as headers
  const keys = Object.keys(data[0]);
  
  // Create header row
  const result: any[][] = [keys];
  
  // Add data rows
  data.forEach(obj => {
    const row = keys.map(key => obj[key]);
    result.push(row);
  });
  
  return result;
}

// Populate table with provided data
async function populateTable(data: any[]): Promise<void> {
  const selection = figma.currentPage.selection;
  
  if (selection.length === 0) {
    figma.notify('Please select table cells or rows to populate');
    return;
  }
  
  // Get all cells from selection
  const allCells = getAllCells(selection);
  
  if (allCells.length === 0) {
    figma.notify('No table cells found in selection');
    return;
  }
  
  // Organize cells into rows
  const rows = organizeCellsIntoRows(allCells);
  
  // Get all text nodes within the cells to load fonts
  const allTextNodes: SceneNode[] = [];
  allCells.forEach(cell => {
    const findTextNodes = (node: SceneNode) => {
      if (node.type === 'TEXT') {
        allTextNodes.push(node);
      } else if ('children' in node) {
        const parent = node as ChildrenMixin & SceneNode;
        parent.children.forEach(child => findTextNodes(child));
      }
    };
    findTextNodes(cell);
  });
  
  await loadFonts(allTextNodes);
  
  // Convert data to 2D array format
  const dataRows = convertObjectsTo2DArray(data);
  
  // Populate cells with data
  let updatedCount = 0;
  
  for (let rowIndex = 0; rowIndex < rows.length && rowIndex < dataRows.length; rowIndex++) {
    const rowCells = rows[rowIndex];
    const rowData = dataRows[rowIndex];
    
    for (let colIndex = 0; colIndex < rowCells.length && colIndex < rowData.length; colIndex++) {
      const cell = rowCells[colIndex];
      const valueLayer = findValueLayer(cell);
      
      if (valueLayer) {
        valueLayer.characters = String(rowData[colIndex] !== null && rowData[colIndex] !== undefined ? rowData[colIndex] : '');
        updatedCount++;
      }
    }
  }
  
  // Success notification with checkmark
  figma.notify(`✅ Updated ${updatedCount} cells with data`);
}

// Generate and populate with dummy data
async function populateTableWithDummyData(rowCount: number, columnCount: number): Promise<void> {
  const selection = figma.currentPage.selection;
  
  if (selection.length === 0) {
    figma.notify('Please select table cells or rows to populate');
    return;
  }
  
  // Get all cells from selection
  const allCells = getAllCells(selection);
  
  if (allCells.length === 0) {
    figma.notify('No table cells found in selection');
    return;
  }
  
  // Organize cells into rows
  const rows = organizeCellsIntoRows(allCells);
  
  // Get all text nodes within the cells to load fonts
  const allTextNodes: SceneNode[] = [];
  allCells.forEach(cell => {
    const findTextNodes = (node: SceneNode) => {
      if (node.type === 'TEXT') {
        allTextNodes.push(node);
      } else if ('children' in node) {
        const parent = node as ChildrenMixin & SceneNode;
        parent.children.forEach(child => findTextNodes(child));
      }
    };
    findTextNodes(cell);
  });
  
  await loadFonts(allTextNodes);
  
  // Generate dummy data
  const departments = [
    "Design", "Engineering", "Marketing", "Product", "Sales", 
    "Finance", "Legal", "Human Resources", "Customer Support", "Operations",
    "Research", "Development", "Quality Assurance", "Business Development", "IT"
  ];
  
  const locations = [
    "London", "New York", "Singapore", "Berlin", "Madrid", 
    "Paris", "Tokyo", "Sydney", "Toronto", "Dubai",
    "Amsterdam", "Stockholm", "Hong Kong", "Mumbai", "São Paulo",
    "Seoul", "Mexico City", "Cape Town", "Milan", "Zurich"
  ];
  
  const accessLevels = ["Admin", "Editor", "Viewer", "Owner", "Guest"];
  const statuses = ["Active", "Inactive", "Pending", "Suspended", "Archived"];
  
  // First names with more variety
  const firstNames = [
    "Mark", "Sarah", "James", "Priya", "Michael", 
    "Emma", "David", "Olivia", "Thomas", "Aisha",
    "Robert", "Sofia", "Daniel", "Mei", "Carlos", 
    "Fatima", "John", "Zara", "Alexander", "Leila",
    "William", "Sophia", "Luis", "Elena", "Mohammed",
    "Chloe", "Hiroshi", "Isabella", "Raj", "Ingrid"
  ];
  
  // Last names with more variety
  const lastNames = [
    "Darnalds", "Johnson", "Chen", "Patel", "Rodriguez", 
    "Wilson", "Kim", "Martinez", "Schmidt", "Ahmed",
    "Brown", "Garcia", "Nguyen", "Singh", "Müller",
    "Taylor", "Sato", "Lopez", "Ivanov", "Silva",
    "Anderson", "Kowalski", "Rossi", "Jensen", "Ali",
    "Wang", "Dubois", "Hernandez", "O'Connor", "Yamamoto"
  ];
  
  // Shuffle arrays for randomization
  const shuffleArray = <T>(array: T[]): T[] => {
    const shuffled = [...array];
    for (let i = shuffled.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    return shuffled;
  };
  
  const shuffledFirstNames = shuffleArray(firstNames);
  const shuffledLastNames = shuffleArray(lastNames);
  const shuffledDepartments = shuffleArray(departments);
  const shuffledLocations = shuffleArray(locations);
  
  // Create dummy data objects
  const dummyData: any[] = [];
  
  for (let i = 0; i < Math.max(rowCount, 30); i++) {
    // Use modulo to cycle through the shuffled arrays
    const firstName = shuffledFirstNames[i % shuffledFirstNames.length];
    const lastName = shuffledLastNames[(i + 3) % shuffledLastNames.length]; // Offset to avoid matching patterns
    
    // Generate a unique email with a small random variation
    const emailPrefix = Math.random() < 0.2 
      ? `${firstName.toLowerCase()[0]}${lastName.toLowerCase()}`
      : `${firstName.toLowerCase()}.${lastName.toLowerCase()}`;
    
    const email = `${emailPrefix}@company.com`;
    
    // Select from shuffled arrays with some randomness
    const department = shuffledDepartments[i % shuffledDepartments.length];
    const location = shuffledLocations[(i + Math.floor(Math.random() * 3)) % shuffledLocations.length];
    const accessLevel = accessLevels[Math.floor(Math.random() * accessLevels.length)];
    
    // Make "Active" status more common than others
    const status = Math.random() < 0.7 ? "Active" : statuses[Math.floor(Math.random() * statuses.length)];
    
    dummyData.push({
      "Name": `${firstName} ${lastName}`,
      "Department": department,
      "Email": email,
      "Location": location,
      "Access Level": accessLevel,
      "Status": status
    });
  }
  
  // Shuffle the final data array for extra randomness
  const shuffledData = shuffleArray(dummyData).slice(0, rowCount);
  
  // Convert to 2D array for populating
  const dataRows = convertObjectsTo2DArray(shuffledData);
  
  // Populate cells with dummy data
  let updatedCount = 0;
  
  for (let rowIndex = 0; rowIndex < rows.length && rowIndex < dataRows.length; rowIndex++) {
    const rowCells = rows[rowIndex];
    const rowData = dataRows[rowIndex];
    
    for (let colIndex = 0; colIndex < rowCells.length && colIndex < rowData.length; colIndex++) {
      const cell = rowCells[colIndex];
      const valueLayer = findValueLayer(cell);
      
      if (valueLayer) {
        valueLayer.characters = String(rowData[colIndex]);
        updatedCount++;
      }
    }
  }
  
  // Success notification with checkmark
  figma.notify(`✅ Updated ${updatedCount} cells with dummy data`);
}
