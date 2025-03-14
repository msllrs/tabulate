// Show the plugin UI
figma.showUI(__html__, { width: 420, height: 540 });

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

type GenerateCustomDummyDataMessage = {
  type: 'generate-custom-dummy-data';
  headers: string[];
  rowCount: number;
};

type CancelMessage = {
  type: 'cancel';
};

type CheckSelectionMessage = {
  type: 'check-selection';
};

type ShowNotificationMessage = {
  type: 'show-notification';
  message: string;
};

type PluginMessage = 
  | PopulateTableMessage 
  | GenerateDummyDataMessage 
  | GenerateCustomDummyDataMessage
  | CancelMessage
  | CheckSelectionMessage
  | ShowNotificationMessage;

// Listen for messages from the UI
figma.ui.onmessage = async (msg: PluginMessage) => {
  if (msg.type === 'populate-table') {
    await populateTable(msg.data);
  } else if (msg.type === 'generate-dummy-data') {
    await populateTableWithDummyData(msg.rowCount, msg.columnCount);
  } else if (msg.type === 'generate-custom-dummy-data') {
    await populateTableWithCustomDummyData(msg.headers, msg.rowCount);
  } else if (msg.type === 'cancel') {
    figma.closePlugin();
  } else if (msg.type === 'check-selection') {
    const hasSelection = figma.currentPage.selection.length > 0;
    
    figma.ui.postMessage({
      type: 'selection-result',
      hasSelection: hasSelection
    });
    
    if (!hasSelection) {
      figma.notify('🚨 Please select table cells or rows to populate');
    }
  } else if (msg.type === 'show-notification') {
    figma.notify(msg.message);
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
    figma.notify('🚨 Please select table cells or rows to populate');
    return;
  }
  
  // Get all cells from selection
  const allCells = getAllCells(selection);
  
  if (allCells.length === 0) {
    figma.notify('🚨 No table cells found in selection');
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
    figma.notify('🚨 Please select table cells or rows to populate');
    return;
  }
  
  // Get all cells from selection
  const allCells = getAllCells(selection);
  
  if (allCells.length === 0) {
    figma.notify('🚨 No table cells found in selection');
    return;
  }
  
  // Organize cells into rows
  const rows = organizeCellsIntoRows(allCells);
  
  // Determine actual column count based on the first row
  const actualColumnCount = rows[0]?.length || 0;
  
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
  
  // Additional data fields for extended columns
  const teams = [
    "Alpha", "Beta", "Gamma", "Delta", "Epsilon", 
    "Omega", "Phoenix", "Titan", "Nexus", "Quantum",
    "Horizon", "Apex", "Zenith", "Pulse", "Fusion"
  ];
  
  const projects = [
    "Dashboard Redesign", "Mobile App", "API Integration", "Data Migration", "Cloud Infrastructure", 
    "Security Audit", "UI Component Library", "Analytics Platform", "Customer Portal", "Automation Framework",
    "Blockchain Solution", "Machine Learning Model", "IoT Platform", "AR Experience", "Payment Gateway"
  ];
  
  const skills = [
    "JavaScript", "Python", "Design Systems", "Product Strategy", "Data Analysis", 
    "Project Management", "UX Research", "Cloud Architecture", "DevOps", "Machine Learning",
    "Blockchain", "Mobile Development", "Cybersecurity", "AI", "Leadership"
  ];
  
  const languages = [
    "English", "Spanish", "Mandarin", "French", "German", 
    "Japanese", "Russian", "Arabic", "Portuguese", "Hindi",
    "Korean", "Italian", "Dutch", "Swedish", "Turkish"
  ];
  
  const startDates = [
    "Jan 2022", "Mar 2021", "Sep 2020", "Feb 2023", "Nov 2019", 
    "Apr 2022", "Jul 2021", "Dec 2020", "May 2023", "Aug 2018",
    "Oct 2022", "Jun 2021", "Jan 2020", "Mar 2023", "Sep 2019"
  ];
  
  const phoneNumbers = [
    "+1 (555) 123-4567", "+44 20 7946 0958", "+65 8765 4321", "+49 30 1234 5678", "+34 91 123 4567", 
    "+33 1 23 45 67 89", "+81 3 1234 5678", "+61 2 1234 5678", "+1 (416) 123-4567", "+971 4 123 4567",
    "+31 20 123 4567", "+46 8 123 45 67", "+852 1234 5678", "+91 22 1234 5678", "+55 11 1234-5678"
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
  const shuffledTeams = shuffleArray(teams);
  const shuffledProjects = shuffleArray(projects);
  const shuffledSkills = shuffleArray(skills);
  const shuffledLanguages = shuffleArray(languages);
  const shuffledStartDates = shuffleArray(startDates);
  const shuffledPhoneNumbers = shuffleArray(phoneNumbers);
  
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
    
    // Additional fields for extended columns
    const team = shuffledTeams[i % shuffledTeams.length];
    const project = shuffledProjects[(i + 2) % shuffledProjects.length];
    const skill = shuffledSkills[(i + 4) % shuffledSkills.length];
    const language = shuffledLanguages[(i + 1) % shuffledLanguages.length];
    const startDate = shuffledStartDates[(i + 5) % shuffledStartDates.length];
    const phoneNumber = shuffledPhoneNumbers[(i + 3) % shuffledPhoneNumbers.length];
    
    // Create base object with standard fields
    const dataObj: any = {
      "Name": `${firstName} ${lastName}`,
      "Department": department,
      "Email": email,
      "Location": location,
      "Access Level": accessLevel,
      "Status": status
    };
    
    // Add extended fields if needed
    if (actualColumnCount > 6) {
      dataObj["Team"] = team;
    }
    if (actualColumnCount > 7) {
      dataObj["Project"] = project;
    }
    if (actualColumnCount > 8) {
      dataObj["Skill"] = skill;
    }
    if (actualColumnCount > 9) {
      dataObj["Language"] = language;
    }
    if (actualColumnCount > 10) {
      dataObj["Start Date"] = startDate;
    }
    if (actualColumnCount > 11) {
      dataObj["Phone"] = phoneNumber;
    }
    
    dummyData.push(dataObj);
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

// Generate and populate with custom dummy data
async function populateTableWithCustomDummyData(headers: string[], rowCount: number): Promise<void> {
  const selection = figma.currentPage.selection;
  
  if (selection.length === 0) {
    figma.notify('🚨 Please select table cells or rows to populate');
    return;
  }
  
  // Get all cells from selection
  const allCells = getAllCells(selection);
  
  if (allCells.length === 0) {
    figma.notify('🚨 No table cells found in selection');
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
  
  // Generate data based on headers
  const data = generateDataFromHeaders(headers, rowCount);
  
  // Convert to 2D array for populating
  const dataRows = [headers, ...data]; // Add headers as first row
  
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
  figma.notify(`✅ Updated ${updatedCount} cells with custom data`);
}

// Function to generate data based on header names
function generateDataFromHeaders(headers: string[], rowCount: number): string[][] {
  // Data generation resources
  const firstNames = [
    "Mark", "Sarah", "James", "Priya", "Michael", 
    "Emma", "David", "Olivia", "Thomas", "Aisha",
    "Robert", "Sofia", "Daniel", "Mei", "Carlos", 
    "Fatima", "John", "Zara", "Alexander", "Leila",
    "William", "Sophia", "Luis", "Elena", "Mohammed",
    "Chloe", "Hiroshi", "Isabella", "Raj", "Ingrid"
  ];
  
  const lastNames = [
    "Darnalds", "Johnson", "Chen", "Patel", "Rodriguez", 
    "Wilson", "Kim", "Martinez", "Schmidt", "Ahmed",
    "Brown", "Garcia", "Nguyen", "Singh", "Müller",
    "Taylor", "Sato", "Lopez", "Ivanov", "Silva",
    "Anderson", "Kowalski", "Rossi", "Jensen", "Ali",
    "Wang", "Dubois", "Hernandez", "O'Connor", "Yamamoto"
  ];
  
  const departments = [
    "Design", "Engineering", "Marketing", "Product", "Sales", 
    "Finance", "Legal", "Human Resources", "Customer Support", "Operations",
    "Research", "Development", "Quality Assurance", "Business Development", "IT"
  ];
  
  const locations = [
    "London", "New York", "Singapore", "Berlin", "Madrid", 
    "Paris", "Tokyo", "Sydney", "Toronto", "Dubai",
    "Amsterdam", "Stockholm", "Hong Kong", "Mumbai", "São Paulo"
  ];
  
  const roles = [
    "Manager", "Director", "Associate", "Specialist", "Lead", 
    "Coordinator", "Analyst", "Designer", "Developer", "Consultant",
    "Administrator", "Strategist", "Engineer", "Architect", "Executive"
  ];
  
  const teams = [
    "Alpha", "Beta", "Gamma", "Delta", "Epsilon", 
    "Omega", "Phoenix", "Titan", "Nexus", "Quantum",
    "Horizon", "Apex", "Zenith", "Pulse", "Fusion"
  ];
  
  const projects = [
    "Dashboard Redesign", "Mobile App", "API Integration", "Data Migration", "Cloud Infrastructure", 
    "Security Audit", "UI Component Library", "Analytics Platform", "Customer Portal", "Automation Framework"
  ];
  
  const skills = [
    "JavaScript", "Python", "Design Systems", "Product Strategy", "Data Analysis", 
    "Project Management", "UX Research", "Cloud Architecture", "DevOps", "Machine Learning"
  ];
  
  const languages = [
    "English", "Spanish", "Mandarin", "French", "German", 
    "Japanese", "Russian", "Arabic", "Portuguese", "Hindi"
  ];
  
  const statuses = ["Active", "Inactive", "Pending", "Suspended", "Archived", "Completed", "In Progress", "On Hold"];
  
  const companies = [
    "Acme Inc.", "Globex Corp", "Initech", "Umbrella Corp", "Stark Industries", 
    "Wayne Enterprises", "Cyberdyne Systems", "Soylent Corp", "Massive Dynamic", "Hooli"
  ];
  
  const addresses = [
    "123 Main St", "456 Park Ave", "789 Broadway", "321 Oak Lane", "654 Pine Road", 
    "987 Maple Drive", "741 Cedar Blvd", "852 Elm Street", "963 Willow Way", "159 Birch Court"
  ];
  
  const cities = [
    "New York", "London", "Tokyo", "Paris", "Berlin", 
    "Sydney", "Toronto", "Singapore", "Dubai", "Mumbai"
  ];
  
  const phoneNumbers = [
    "+1 (555) 123-4567", "+44 20 7946 0958", "+65 8765 4321", "+49 30 1234 5678", "+34 91 123 4567", 
    "+33 1 23 45 67 89", "+81 3 1234 5678", "+61 2 1234 5678", "+1 (416) 123-4567", "+971 4 123 4567"
  ];
  
  const dates = [
    "2023-01-15", "2023-02-28", "2023-03-10", "2023-04-22", "2023-05-05", 
    "2023-06-18", "2023-07-30", "2023-08-12", "2023-09-25", "2023-10-08",
    "2023-11-20", "2023-12-03", "2024-01-16", "2024-02-29", "2024-03-11"
  ];
  
  const prices = [
    "$19.99", "$24.50", "$99.00", "$149.99", "$7.99", 
    "$35.75", "$199.99", "$49.00", "$299.99", "$12.50",
    "$59.99", "$85.25", "$129.00", "$14.99", "$249.99"
  ];
  
  const quantities = ["1", "2", "5", "10", "25", "50", "100", "250", "500", "1000"];
  
  const ratings = ["★★★★★", "★★★★☆", "★★★☆☆", "★★☆☆☆", "★☆☆☆☆"];
  
  const ids = [];
  for (let i = 1; i <= 100; i++) {
    // Use a simple string padding approach instead of padStart
    const paddedNum = ('0000' + i).slice(-4);
    ids.push(`ID-${paddedNum}`);
  }  
  
  // Shuffle arrays for randomization
  const shuffleArray = <T>(array: T[]): T[] => {
    const shuffled = [...array];
    for (let i = shuffled.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    return shuffled;
  };
  
  // Shuffle all data arrays
  const shuffledFirstNames = shuffleArray(firstNames);
  const shuffledLastNames = shuffleArray(lastNames);
  const shuffledDepartments = shuffleArray(departments);
  const shuffledLocations = shuffleArray(locations);
  const shuffledRoles = shuffleArray(roles);
  const shuffledTeams = shuffleArray(teams);
  const shuffledProjects = shuffleArray(projects);
  const shuffledSkills = shuffleArray(skills);
  const shuffledLanguages = shuffleArray(languages);
  const shuffledStatuses = shuffleArray(statuses);
  const shuffledCompanies = shuffleArray(companies);
  const shuffledAddresses = shuffleArray(addresses);
  const shuffledCities = shuffleArray(cities);
  const shuffledPhoneNumbers = shuffleArray(phoneNumbers);
  const shuffledDates = shuffleArray(dates);
  const shuffledPrices = shuffleArray(prices);
  const shuffledQuantities = shuffleArray(quantities);
  const shuffledRatings = shuffleArray(ratings);
  const shuffledIds = shuffleArray(ids);
  
  // Function to determine data type from header name
  function getDataTypeForHeader(header: string): string {
    const headerLower = header.toLowerCase().trim();
    
    // Name-related headers
    if (headerLower.includes('first') && (headerLower.includes('name') || headerLower === 'first')) {
      return 'firstName';
    } else if (headerLower.includes('last') && (headerLower.includes('name') || headerLower === 'last')) {
      return 'lastName';
    } else if (headerLower === 'name' || headerLower === 'full name' || headerLower.includes('full') && headerLower.includes('name')) {
      return 'fullName';
    }
    
    // Contact information
    else if (headerLower.includes('email') || headerLower.includes('e-mail') || headerLower.includes('mail')) {
      return 'email';
    } else if (headerLower.includes('phone') || headerLower.includes('mobile') || headerLower.includes('cell') || 
              (headerLower.includes('number') && !headerLower.includes('id'))) {
      return 'phone';
    } else if (headerLower.includes('website') || headerLower.includes('site') || headerLower.includes('url') || 
              headerLower.includes('web') || headerLower.includes('link')) {
      return 'website';
    }
    
    // Location-related headers
    else if (headerLower.includes('address') || headerLower.includes('street')) {
      return 'address';
    } else if (headerLower === 'city') {
      return 'city';
    } else if (headerLower === 'state' || headerLower === 'province') {
      return 'state';
    } else if (headerLower === 'country') {
      return 'country';
    } else if (headerLower.includes('zip') || headerLower.includes('postal') || headerLower.includes('post code')) {
      return 'zipCode';
    } else if (headerLower.includes('location') || headerLower.includes('region') || headerLower.includes('area')) {
      return 'location';
    }
    
    // Organizational headers
    else if (headerLower.includes('department') || headerLower.includes('dept')) {
      return 'department';
    } else if (headerLower.includes('role') || headerLower.includes('position') || headerLower.includes('title') || 
              headerLower.includes('job')) {
      return 'role';
    } else if (headerLower.includes('team') || headerLower.includes('group')) {
      return 'team';
    } else if (headerLower.includes('project') || headerLower.includes('initiative')) {
      return 'project';
    } else if (headerLower.includes('company') || headerLower.includes('organization') || 
              headerLower.includes('employer') || headerLower.includes('business')) {
      return 'company';
    }
    
    // Skills and attributes
    else if (headerLower.includes('skill') || headerLower.includes('expertise') || 
            headerLower.includes('specialty') || headerLower.includes('proficiency')) {
      return 'skill';
    } else if (headerLower.includes('language') || headerLower.includes('lang')) {
      return 'language';
    }
    
    // Status and metrics
    else if (headerLower === 'status' || headerLower.includes('state') && !headerLower.includes('country')) {
      return 'status';
    } else if (headerLower.includes('date') || headerLower.includes('time') || headerLower.includes('when')) {
      return 'date';
    } else if (headerLower.includes('price') || headerLower.includes('cost') || headerLower.includes('$') || 
              headerLower.includes('fee') || headerLower.includes('amount') && !headerLower.includes('quantity')) {
      return 'price';
    } else if (headerLower.includes('quantity') || headerLower.includes('qty') || 
              headerLower.includes('count') || headerLower.includes('number of')) {
      return 'quantity';
    } else if (headerLower.includes('rating') || headerLower.includes('score') || 
              headerLower.includes('rank') || headerLower.includes('review')) {
      return 'rating';
    }
    
    // Identifiers
    else if (headerLower === 'id' || headerLower.includes('identifier') || headerLower.includes('uuid') || 
            headerLower.includes('code') || headerLower.includes('number') && headerLower.includes('id')) {
      return 'id';
    } else if (headerLower.includes('access') || headerLower.includes('permission') || headerLower.includes('privilege')) {
      return 'accessLevel';
    }
    
    // Descriptions and content
    else if (headerLower.includes('description') || headerLower.includes('desc') || 
            headerLower.includes('about') || headerLower.includes('details')) {
      return 'description';
    } else if (headerLower.includes('comment') || headerLower.includes('note') || headerLower.includes('feedback')) {
      return 'comment';
    }
    
    // Time-related
    else if (headerLower.includes('duration') || headerLower.includes('length') || headerLower.includes('time span')) {
      return 'duration';
    } else if (headerLower.includes('start') && headerLower.includes('date')) {
      return 'startDate';
    } else if (headerLower.includes('end') && headerLower.includes('date')) {
      return 'endDate';
    } else if (headerLower.includes('deadline')) {
      return 'deadline';
    }
    
    // Financial
    else if (headerLower.includes('budget')) {
      return 'budget';
    } else if (headerLower.includes('revenue') || headerLower.includes('sales')) {
      return 'revenue';
    } else if (headerLower.includes('profit') || headerLower.includes('margin')) {
      return 'profit';
    }
    
    // For any other header, use a generic type based on the header name itself
    return 'generic:' + header.trim();
  }
  
  // Function to generate data for a specific type and row index
  function generateDataForType(type: string, rowIndex: number): string {
    // Handle generic types (headers that don't match any pattern)
    if (type.startsWith('generic:')) {
      const headerName = type.substring(8); // Remove 'generic:' prefix
      
      // Generate sensible data based on the header name itself
      // This ensures we always return something reasonable even for unknown headers
      
      // Check if the header might be a percentage
      if (headerName.toLowerCase().includes('percent') || headerName.toLowerCase().includes('%')) {
        return `${Math.floor(Math.random() * 100)}%`;
      }
      
      // Check if the header might be a year
      if (headerName.toLowerCase().includes('year')) {
        const currentYear = new Date().getFullYear();
        return `${currentYear - Math.floor(Math.random() * 5)}`;
      }
      
      // Check if the header might be related to currency but wasn't caught earlier
      if (headerName.toLowerCase().includes('price') || headerName.toLowerCase().includes('cost') || 
          headerName.toLowerCase().includes('payment') || headerName.toLowerCase().includes('salary')) {
        return `$${(Math.random() * 1000).toFixed(2)}`;
      }
      
      // For truly generic headers, generate simple text with the row number
      return `${headerName} ${rowIndex + 1}`;
    }
    
    // Handle all the specific data types
    switch (type) {
      case 'fullName':
        return `${shuffledFirstNames[rowIndex % shuffledFirstNames.length]} ${shuffledLastNames[rowIndex % shuffledLastNames.length]}`;
      case 'firstName':
        return shuffledFirstNames[rowIndex % shuffledFirstNames.length];
      case 'lastName':
        return shuffledLastNames[rowIndex % shuffledLastNames.length];
      case 'email': {
        const firstName = shuffledFirstNames[rowIndex % shuffledFirstNames.length];
        const lastName = shuffledLastNames[rowIndex % shuffledLastNames.length];
        return `${firstName.toLowerCase()}.${lastName.toLowerCase()}@company.com`;
      }
      case 'phone':
        return shuffledPhoneNumbers[rowIndex % shuffledPhoneNumbers.length];
      case 'website': {
        const company = shuffledCompanies[rowIndex % shuffledCompanies.length].toLowerCase().replace(/[^a-z0-9]/g, '');
        return `https://www.${company}.com`;
      }
      case 'department':
        return shuffledDepartments[rowIndex % shuffledDepartments.length];
      case 'location':
        return shuffledLocations[rowIndex % shuffledLocations.length];
      case 'city':
        return shuffledCities[rowIndex % shuffledCities.length];
      case 'state':
        return ["California", "New York", "Texas", "Florida", "Illinois", "Pennsylvania", "Ohio", "Georgia", "Michigan", "North Carolina"][rowIndex % 10];
      case 'country':
        return ["United States", "United Kingdom", "Canada", "Australia", "Germany", "France", "Japan", "India", "Brazil", "China"][rowIndex % 10];
      case 'zipCode':
        return `${10000 + Math.floor(Math.random() * 90000)}`;
      case 'role':
        return shuffledRoles[rowIndex % shuffledRoles.length];
      case 'team':
        return shuffledTeams[rowIndex % shuffledTeams.length];
      case 'project':
        return shuffledProjects[rowIndex % shuffledProjects.length];
      case 'skill':
        return shuffledSkills[rowIndex % shuffledSkills.length];
      case 'language':
        return shuffledLanguages[rowIndex % shuffledLanguages.length];
      case 'status':
        return shuffledStatuses[rowIndex % shuffledStatuses.length];
      case 'company':
        return shuffledCompanies[rowIndex % shuffledCompanies.length];
      case 'address':
        return `${shuffledAddresses[rowIndex % shuffledAddresses.length]}, ${shuffledCities[rowIndex % shuffledCities.length]}`;
      case 'date':
        return shuffledDates[rowIndex % shuffledDates.length];
      case 'startDate':
        return shuffledDates[(rowIndex + 2) % shuffledDates.length];
      case 'endDate':
        return shuffledDates[(rowIndex + 5) % shuffledDates.length];
      case 'deadline':
        // Generate a future date
        const futureDate = new Date();
        futureDate.setDate(futureDate.getDate() + (10 + rowIndex % 30));
        return futureDate.toISOString().split('T')[0];
      case 'price':
        return shuffledPrices[rowIndex % shuffledPrices.length];
      case 'budget':
        return `$${(10000 + Math.floor(Math.random() * 90000)).toLocaleString()}`;
      case 'revenue':
        return `$${(100000 + Math.floor(Math.random() * 900000)).toLocaleString()}`;
      case 'profit':
        return `$${(5000 + Math.floor(Math.random() * 50000)).toLocaleString()}`;
      case 'quantity':
        return shuffledQuantities[rowIndex % shuffledQuantities.length];
      case 'rating':
        return shuffledRatings[rowIndex % shuffledRatings.length];
      case 'id':
        return shuffledIds[rowIndex % shuffledIds.length];
      case 'accessLevel':
        return ['Admin', 'Editor', 'Viewer', 'Owner', 'Guest'][rowIndex % 5];
      case 'description':
        return [
          "A comprehensive solution for enterprise needs.",
          "Designed with user experience in mind.",
          "Optimized for performance and scalability.",
          "Innovative approach to solving complex problems.",
          "Industry-leading features and capabilities.",
          "Built on cutting-edge technology.",
          "Tailored for maximum efficiency.",
          "Seamless integration with existing systems.",
          "Revolutionary design that sets new standards.",
          "Engineered for reliability and durability."
        ][rowIndex % 10];
      case 'comment':
        return [
          "Looks great, approved!",
          "Need some minor adjustments.",
          "Excellent work on this.",
          "Let's discuss this further.",
          "I have some concerns about this.",
          "Very impressive results.",
          "This exceeds expectations.",
          "Please review and get back to me.",
          "I'd like to see alternative options.",
          "This is exactly what we needed."
        ][rowIndex % 10];
      case 'duration':
        return [`${1 + rowIndex % 12} hours`, `${1 + rowIndex % 30} days`, `${1 + rowIndex % 12} months`][rowIndex % 3];
      default:
        return `Item ${rowIndex + 1}`;
    }
  }
  
  // Generate data rows
  const rows: string[][] = [];
  
  // Determine data type for each header
  const headerTypes = headers.map(header => getDataTypeForHeader(header));
  
  // Generate data for each row
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    const row: string[] = [];
    
    // Generate data for each column based on its header type
    for (let colIndex = 0; colIndex < headers.length; colIndex++) {
      const dataType = headerTypes[colIndex];
      row.push(generateDataForType(dataType, rowIndex));
    }
    
    rows.push(row);
  }
  
  return rows;
}
