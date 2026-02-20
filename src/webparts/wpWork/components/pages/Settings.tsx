/**
 * Settings.tsx
 * Settings management page component
 * Configure application preferences and options
 */

import * as React from 'react';
import {
  Stack,
  Text,
  TextField,
  Toggle,
  PrimaryButton,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import styles from '../AppBoilerplate.module.scss';
import { SPService } from '../../services/SPService';
import { IPageComponentProps } from './IPageComponentProps';

export interface ISettingsProps extends IPageComponentProps {}

export const Settings: React.FC<ISettingsProps> = ({ spService }) => {
  const [settings, setSettings] = React.useState({
    itemsPerPage: 10,
    enableNotifications: true,
    autoRefresh: false,
    refreshInterval: 30,
  });

  const [saved, setSaved] = React.useState(false);
  const [loading, setLoading] = React.useState(false);

  const handleSaveSettings = async () => {
    try {
      setLoading(true);
      // Simulate saving settings
      await new Promise((resolve) => setTimeout(resolve, 500));
      setSaved(true);
      setTimeout(() => setSaved(false), 3000);
    } catch (error) {
      console.error('Error saving settings:', error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <Stack className={styles.pageContent} tokens={{ childrenGap: 20 }}>
      <Stack.Item>
        <Text variant="xLarge" block>
          Settings
        </Text>
        <Text variant="small" className={styles.subtitle}>
          Configure application preferences
        </Text>
      </Stack.Item>

      {saved && (
        <Stack.Item>
          <MessageBar messageBarType={MessageBarType.success}>
            Settings saved successfully!
          </MessageBar>
        </Stack.Item>
      )}

      <Stack.Item>
        <Stack className={styles.metricCard} tokens={{ childrenGap: 20 }}>
          <Stack.Item>
            <Text variant="large" className={styles.sectionTitle}>
              Display Settings
            </Text>
          </Stack.Item>

          <Stack.Item>
            <TextField
              label="Items Per Page"
              type="number"
              min={1}
              max={100}
              value={settings.itemsPerPage.toString()}
              onChange={(e, value) =>
                setSettings({
                  ...settings,
                  itemsPerPage: parseInt(value || '10', 10),
                })
              }
            />
          </Stack.Item>

          <Stack.Item>
            <Toggle
              label="Enable Notifications"
              checked={settings.enableNotifications}
              onChange={(e, checked) =>
                setSettings({
                  ...settings,
                  enableNotifications: checked || false,
                })
              }
              onText="On"
              offText="Off"
            />
          </Stack.Item>
        </Stack>
      </Stack.Item>

      <Stack.Item>
        <Stack className={styles.metricCard} tokens={{ childrenGap: 20 }}>
          <Stack.Item>
            <Text variant="large" className={styles.sectionTitle}>
              Refresh Settings
            </Text>
          </Stack.Item>

          <Stack.Item>
            <Toggle
              label="Auto-Refresh Data"
              checked={settings.autoRefresh}
              onChange={(e, checked) =>
                setSettings({
                  ...settings,
                  autoRefresh: checked || false,
                })
              }
              onText="On"
              offText="Off"
            />
          </Stack.Item>

          {settings.autoRefresh && (
            <Stack.Item>
              <TextField
                label="Refresh Interval (seconds)"
                type="number"
                min={5}
                max={300}
                value={settings.refreshInterval.toString()}
                onChange={(e, value) =>
                  setSettings({
                    ...settings,
                    refreshInterval: parseInt(value || '30', 10),
                  })
                }
              />
            </Stack.Item>
          )}
        </Stack>
      </Stack.Item>

      <Stack.Item>
        <Stack className={styles.metricCard} tokens={{ childrenGap: 20 }}>
          <Stack.Item>
            <Text variant="large" className={styles.sectionTitle}>
              About This Application
            </Text>
          </Stack.Item>

          <Stack.Item>
            <Stack tokens={{ childrenGap: 8 }}>
              <Text>
                <strong>Application:</strong> Single Part App Page
              </Text>
              <Text>
                <strong>Version:</strong> 1.0.0
              </Text>
              <Text>
                <strong>Framework:</strong> SharePoint Framework (SPFx) v1.22+
              </Text>
              <Text>
                <strong>Tech Stack:</strong> React 17, Fluent UI, PnP.js
              </Text>
            </Stack>
          </Stack.Item>
        </Stack>
      </Stack.Item>

      <Stack.Item>
        <PrimaryButton
          text="Save Settings"
          onClick={handleSaveSettings}
          disabled={loading}
        />
      </Stack.Item>
    </Stack>
  );
};

export default Settings;
