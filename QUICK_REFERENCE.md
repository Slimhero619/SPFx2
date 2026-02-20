/**
 * QUICK_REFERENCE.md
 * Quick code examples for using the SPFx Single Part App Page boilerplate
 */

# SPFx Single Part App Page - Quick Reference Guide

## Manifest Configuration Snippet

To enable **SharePointFullPage** support, ensure your `WpWorkWebPart.manifest.json` contains:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json",
  "supportedHosts": [
    "SharePointWebPart",
    "TeamsPersonalApp", 
    "TeamsTab",
    "SharePointFullPage"  // ← KEY: Enables full-page immersive experience
  ],
  "supportsThemeVariants": true,
  "preconfiguredEntries": [{
    "groupId": "5c03119e-3074-46fd-976b-c60198311f70",
    "group": { "default": "Advanced" },
    "title": { "default": "WpWork" },
    "description": { "default": "Modern Single Part App Page" },
    "officeFabricIconFontName": "Page",
    "properties": {}
  }]
}
```

---

## SPService.ts - PnP.js Service Layer

### Location
```
src/webparts/wpWork/services/SPService.ts
```

### Initialization with SPFx Context

```typescript
import { SPService } from './services/SPService';
import { IWebPartContext } from '@microsoft/sp-webpart-base';

// In your component
const spService = React.useMemo(
  () => new SPService(context),
  [context]
);
```

### Example: Fetch Data from VATaskList

```typescript
const [tasks, setTasks] = React.useState<ITaskItem[]>([]);
const [loading, setLoading] = React.useState(true);
const [error, setError] = React.useState<string | null>(null);

React.useEffect(() => {
  const fetchTasks = async () => {
    try {
      setLoading(true);
      const fetchedTasks = await spService.getTasksFromVATaskList();
      setTasks(fetchedTasks);
    } catch (err) {
      setError('Failed to load tasks');
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  fetchTasks();
}, [spService]);
```

### SPService Methods Reference

```typescript
// Get all tasks
const tasks = await spService.getTasksFromVATaskList();
// Returns: ITaskItem[]

// Get single task
const task = await spService.getTaskById(1);
// Returns: ITaskItem

// Create task
await spService.createTask({
  Title: 'New Task',
  Status: 'Pending',
  DueDate: '2026-03-15'
});

// Update task
await spService.updateTask(1, {
  Status: 'Completed'
});

// Delete task
await spService.deleteTask(1);

// Get current user
const user = await spService.getCurrentUser();
```

---

## AppBoilerplate.tsx - Main Component with Routing

### Location
```
src/webparts/wpWork/components/AppBoilerplate.tsx
```

### Component Structure

```typescript
<AppBoilerplate
  context={webPartContext}
  isDarkTheme={isDarkTheme}
  userDisplayName={userDisplayName}
/>
```

### Internal Navigation Flow

1. **HashRouter** wraps everything for client-side routing
2. **Pivot Control** displays tabs: Dashboard, Tasks, Settings
3. **onLinkClick** event triggers `useNavigate()` to change route
4. **Routes** render appropriate page component without full reload

### Routing Example

```typescript
// Navigation state is in URL hash
// /#/dashboard → shows Dashboard
// /#/tasks → shows Tasks  
// /#/settings → shows Settings

const navigate = useNavigate();

// Programmatic navigation
navigate('/tasks');

// From Pivot click
const handlePivotLinkClick = (item?: PivotItem) => {
  if (item?.props.itemKey) {
    navigate(`/${item.props.itemKey}`);
  }
};
```

---

## Page Component Props

### IPageComponentProps Interface

```typescript
export interface IPageComponentProps {
  spService: SPService;
}
```

### Using in Your Page Component

```typescript
import { IPageComponentProps } from './IPageComponentProps';
import { SPService } from '../services/SPService';

export const MyPage: React.FC<IPageComponentProps> = ({ spService }) => {
  // Use spService methods here
  const [data, setData] = React.useState([]);

  React.useEffect(() => {
    const fetchData = async () => {
      const items = await spService.getTasksFromVATaskList();
      setData(items);
    };
    fetchData();
  }, [spService]);

  return <div>{/* Your JSX */}</div>;
};
```

---

## Dashboard Component Example

### Location
```
src/webparts/wpWork/components/pages/Dashboard.tsx
```

### Features

```typescript
// Fetch tasks on component mount
React.useEffect(() => {
  const fetchTasks = async () => {
    try {
      const tasks = await spService.getTasksFromVATaskList();
      setTasks(tasks);
    } catch (err) {
      setError('Failed to load tasks');
    }
  };
  fetchTasks();
}, [spService]);

// Calculate statistics
const completedTasks = tasks.filter(t => t.Status === 'Completed').length;
const pendingTasks = tasks.filter(t => t.Status === 'Pending').length;

// Display in cards
<Card>
  <Text variant="superLarge">{tasks.length}</Text>
  <Text>Total Tasks</Text>
</Card>
```

---

## Tasks Component Example

### Location
```
src/webparts/wpWork/components/pages/Tasks.tsx
```

### Create New Task Dialog

```typescript
const [showNewTaskDialog, setShowNewTaskDialog] = React.useState(false);
const [newTaskTitle, setNewTaskTitle] = React.useState('');
const [newTaskStatus, setNewTaskStatus] = React.useState('Pending');
const [newTaskDueDate, setNewTaskDueDate] = React.useState('');

const handleAddTask = async () => {
  if (!newTaskTitle.trim()) {
    setError('Task title is required');
    return;
  }

  try {
    await spService.createTask({
      Title: newTaskTitle,
      Status: newTaskStatus,
      DueDate: newTaskDueDate || undefined,
    });
    // Clear form and refresh
    setShowNewTaskDialog(false);
    fetchTasks(); // Refetch all tasks
  } catch (err) {
    setError('Failed to create task');
  }
};

// Dialog UI
<Dialog hidden={!showNewTaskDialog} onDismiss={() => setShowNewTaskDialog(false)}>
  <TextField
    label="Task Title"
    value={newTaskTitle}
    onChange={(e, value) => setNewTaskTitle(value || '')}
    required
  />
  <Dropdown
    label="Status"
    options={[
      { key: 'Pending', text: 'Pending' },
      { key: 'In Progress', text: 'In Progress' },
      { key: 'Completed', text: 'Completed' }
    ]}
    selectedKey={newTaskStatus}
    onChange={(e, option) => setNewTaskStatus(option?.key as string)}
  />
  <DialogFooter>
    <PrimaryButton onClick={handleAddTask} text="Create" />
  </DialogFooter>
</Dialog>
```

---

## Settings Component Example

### Location
```
src/webparts/wpWork/components/pages/Settings.tsx
```

### Settings State Management

```typescript
const [settings, setSettings] = React.useState({
  itemsPerPage: 10,
  enableNotifications: true,
  autoRefresh: false,
  refreshInterval: 30,
});

const handleSaveSettings = async () => {
  try {
    setLoading(true);
    // TODO: Save to SharePoint list
    await new Promise((resolve) => setTimeout(resolve, 500));
    setSaved(true);
  } catch (error) {
    console.error('Error saving settings:', error);
  }
};

// Toggle control
<Toggle
  label="Auto-Refresh Data"
  checked={settings.autoRefresh}
  onChange={(e, checked) =>
    setSettings({ ...settings, autoRefresh: checked || false })
  }
/>
```

---

## Styling Components

### AppBoilerplate.module.scss Location
```
src/webparts/wpWork/components/AppBoilerplate.module.scss
```

### Key CSS Classes

```scss
.appContainer        // Main flex container
.navigationHeader    // Pivot header area
.pivotNav           // Pivot styling
.pageContainer      // Scrollable content area
.pageContent        // Page padding and constraints
.cardTitle          // Card heading style
.cardValue          // Large emphasis values
.successColor       // Green for completed
.warningColor       // Orange for pending
.statusCompleted    // Completed item styling
.statusPending      // Pending item styling
.sectionTitle       // Section heading in settings
```

### Custom Styling Example

```scss
.customCard {
  padding: 16px;
  border-radius: 4px;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  
  @media (max-width: 768px) {
    padding: 12px;
  }
}
```

---

## Adding New Page Components

### Step 1: Create Page Component

```typescript
// src/webparts/wpWork/components/pages/Reports.tsx
import React from 'react';
import { Stack, Text } from '@fluentui/react';
import styles from '../AppBoilerplate.module.scss';
import { IPageComponentProps } from './IPageComponentProps';

export const Reports: React.FC<IPageComponentProps> = ({ spService }) => {
  return (
    <Stack className={styles.pageContent}>
      <Text variant="xLarge">Reports</Text>
      {/* Your content here */}
    </Stack>
  );
};

export default Reports;
```

### Step 2: Add Route in AppBoilerplate.tsx

```typescript
import Reports from './pages/Reports';

// In Routes section:
<Route
  path={`/reports`}
  element={<Reports spService={spService} />}
/>
```

### Step 3: Add Pivot Tab

```typescript
<PivotItem
  headerText="Reports"
  itemKey="reports"
/>
```

---

## Extending SPService with Custom Methods

### Add New Method to SPService.ts

```typescript
/**
 * Get tasks by status
 * @param status - Filter by status value
 * @returns Promise of filtered tasks
 */
public async getTasksByStatus(status: string): Promise<ITaskItem[]> {
  try {
    const items = await this.sp.web.lists
      .getByTitle('VATaskList')
      .items
      .filter(`Status eq '${status}'`)
      .select('Id', 'Title', 'Status', 'DueDate')
      ();
    
    return items as ITaskItem[];
  } catch (error) {
    console.error(`Error fetching tasks with status ${status}:`, error);
    throw error;
  }
}
```

### Use in Component

```typescript
const completedTasks = await spService.getTasksByStatus('Completed');
```

---

## Error Handling Pattern

```typescript
try {
  setLoading(true);
  setError(null);
  
  // Perform operation
  const result = await spService.getTasksFromVATaskList();
  setData(result);
  
} catch (err) {
  setError('A user-friendly error message');
  console.error('Technical error details:', err);
  
} finally {
  setLoading(false);
}

// Display to user
{error && (
  <MessageBar messageBarType={MessageBarType.error}>
    {error}
  </MessageBar>
)}
```

---

## Package.json Dependencies

```json
{
  "dependencies": {
    "react": "17.0.1",
    "react-dom": "17.0.1",
    "react-router-dom": "^6.31.0",    // NEW: Client-side routing
    "@pnp/sp": "^4.12.0",              // NEW: SharePoint data access
    "@fluentui/react": "^8.106.4",
    "@microsoft/sp-webpart-base": "1.22.2"
  }
}
```

---

## Build & Deploy Commands

```bash
# Install dependencies
npm install

# Start development server with hot reload
npm start

# Build for production
npm run build

# Package solution
heft test --clean --production && heft package-solution --production

# Clean build artifacts
npm run clean
```

---

## Troubleshooting Checklist

- [ ] Manifest has "SharePointFullPage" in supportedHosts
- [ ] HashRouter wraps AppContent component
- [ ] SPFxTokenProvider is used for PnP.js initialization
- [ ] Page components extend IPageComponentProps interface
- [ ] Routes use hash format (/#/dashboard not /dashboard)
- [ ] SCSS files use .module.scss extension
- [ ] useEffect dependencies array includes [spService]
- [ ] Error states are handled with try-catch blocks
- [ ] Loading states prevent multiple simultaneous requests
- [ ] TypeScript compilation has no errors

---

## File Paths Reference

| File | Purpose |
|------|---------|
| `src/webparts/wpWork/WpWorkWebPart.ts` | Main web part class |
| `src/webparts/wpWork/WpWorkWebPart.manifest.json` | Web part configuration |
| `src/webparts/wpWork/components/WpWork.tsx` | Root component |
| `src/webparts/wpWork/components/AppBoilerplate.tsx` | App with routing |
| `src/webparts/wpWork/components/IWpWorkProps.ts` | Root props interface |
| `src/webparts/wpWork/components/IAppBoilerplateProps.ts` | App props interface |
| `src/webparts/wpWork/components/AppBoilerplate.module.scss` | Main styles |
| `src/webparts/wpWork/components/pages/Dashboard.tsx` | Dashboard page |
| `src/webparts/wpWork/components/pages/Tasks.tsx` | Tasks page |
| `src/webparts/wpWork/components/pages/Settings.tsx` | Settings page |
| `src/webparts/wpWork/services/SPService.ts` | PnP.js service |

---

**Ready to get started!** 🚀

Clone, customize, and deploy to SharePoint as a Single Part App Page.
