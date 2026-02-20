/**
 * AppBoilerplate.tsx
 * Main Single Part App Page component with React Router and Fluent UI navigation
 * Uses HashRouter for client-side routing and Pivot control for tab navigation
 */

import * as React from 'react';
import { HashRouter, Routes, Route, useNavigate, useLocation } from 'react-router-dom';
import {
  Stack,
  Pivot,
  PivotItem,
  ThemeProvider,
  createTheme,
} from '@fluentui/react';
import styles from './AppBoilerplate.module.scss';
import { SPService } from '../services/SPService';
import Dashboard from './pages/Dashboard';
import Tasks from './pages/Tasks';
import Settings from './pages/Settings';
import { IAppBoilerplateProps } from './IAppBoilerplateProps';

// Define navigation items with their routes
export enum NavigationTab {
  Dashboard = 'dashboard',
  Tasks = 'tasks',
  Settings = 'settings',
}

/**
 * Internal routing component (must be inside HashRouter)
 */
const AppContent: React.FC<{ spService: SPService }> = ({ spService }) => {
  const navigate = useNavigate();
  const location = useLocation();

  // Extract current tab from hash route
  const getCurrentTab = (): NavigationTab => {
    const hash = location.pathname.substring(1); // Remove leading /
    return (hash as NavigationTab) || NavigationTab.Dashboard;
  };

  const handlePivotLinkClick = (item?: PivotItem) => {
    if (item?.props.itemKey) {
      navigate(`/${item.props.itemKey}`);
    }
  };

  const currentTab = getCurrentTab();

  return (
    <Stack className={styles.appContainer} tokens={{ childrenGap: 0 }}>
      {/* Navigation Header with Pivot Control */}
      <Stack.Item className={styles.navigationHeader}>
        <Pivot
          selectedKey={currentTab}
          onLinkClick={handlePivotLinkClick}
          headersOnly
          className={styles.pivotNav}
        >
          <PivotItem
            headerText="Dashboard"
            itemKey={NavigationTab.Dashboard}
          />
          <PivotItem
            headerText="Tasks"
            itemKey={NavigationTab.Tasks}
          />
          <PivotItem
            headerText="Settings"
            itemKey={NavigationTab.Settings}
          />
        </Pivot>
      </Stack.Item>

      {/* Page Content Area */}
      <Stack.Item grow className={styles.pageContainer}>
        <Routes>
          <Route
            path={`/${NavigationTab.Dashboard}`}
            element={<Dashboard spService={spService} />}
          />
          <Route
            path={`/${NavigationTab.Tasks}`}
            element={<Tasks spService={spService} />}
          />
          <Route
            path={`/${NavigationTab.Settings}`}
            element={<Settings spService={spService} />}
          />
          {/* Default route */}
          <Route
            path="*"
            element={<Dashboard spService={spService} />}
          />
        </Routes>
      </Stack.Item>
    </Stack>
  );
};

/**
 * Main AppBoilerplate component
 * Wraps everything in HashRouter to enable client-side routing
 */
export const AppBoilerplate: React.FC<IAppBoilerplateProps> = ({
  context,
  isDarkTheme,
  userDisplayName,
}) => {
  const spService = React.useMemo(() => new SPService(context), [context]);

  // Optional: Create a custom theme based on isDarkTheme
  const theme = React.useMemo(
    () =>
      createTheme({
        isInverted: isDarkTheme,
      }),
    [isDarkTheme]
  );

  return (
    <ThemeProvider theme={theme}>
      <HashRouter>
        <AppContent spService={spService} />
      </HashRouter>
    </ThemeProvider>
  );
};

export default AppBoilerplate;
