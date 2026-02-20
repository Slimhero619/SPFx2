/**
 * Dashboard.tsx
 * Dashboard page component for the Single Part App Page
 * Displays overview and key metrics
 */

import * as React from 'react';
import {
  Stack,
  Text,
  MessageBar,
  MessageBarType,
  Shimmer,
} from '@fluentui/react';
import styles from '../AppBoilerplate.module.scss';
import { SPService, ITaskItem } from '../../services/SPService';
import { IPageComponentProps } from './IPageComponentProps';

export interface IDashboardProps extends IPageComponentProps {}

export const Dashboard: React.FC<IDashboardProps> = ({ spService }) => {
  const [tasks, setTasks] = React.useState<ITaskItem[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    const fetchTasks = async () => {
      try {
        setLoading(true);
        setError(null);
        const fetchedTasks = await spService.getTasksFromVATaskList();
        setTasks(fetchedTasks);
      } catch (err) {
        setError('Failed to load tasks. Please try again later.');
        console.error(err);
      } finally {
        setLoading(false);
      }
    };

    fetchTasks();
  }, [spService]);

  const completedTasks = tasks.filter((t) => t.Status === 'Completed').length;
  const pendingTasks = tasks.filter((t) => t.Status === 'Pending').length;

  return (
    <Stack className={styles.pageContent} tokens={{ childrenGap: 20 }}>
      <Stack.Item>
        <Text variant="xLarge" block>
          Dashboard
        </Text>
        <Text variant="small" className={styles.subtitle}>
          Overview of your tasks and progress
        </Text>
      </Stack.Item>

      {error && (
        <Stack.Item>
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        </Stack.Item>
      )}

      {loading ? (
        <Stack.Item>
          <MessageBar messageBarType={MessageBarType.info}>
            Loading data...
          </MessageBar>
        </Stack.Item>
      ) : (
        <>
          <Stack horizontal tokens={{ childrenGap: 20 }}>
            <Stack.Item grow className={styles.metricCard}>
              <Stack tokens={{ childrenGap: 8 }}>
                <Text variant="large" className={styles.cardTitle}>
                  Total Tasks
                </Text>
                <Text variant="superLarge" className={styles.cardValue}>
                  {tasks.length}
                </Text>
              </Stack>
            </Stack.Item>

            <Stack.Item grow className={styles.metricCard}>
              <Stack tokens={{ childrenGap: 8 }}>
                <Text variant="large" className={styles.cardTitle}>
                  Completed
                </Text>
                <Text
                  variant="superLarge"
                  className={`${styles.cardValue} ${styles.successColor}`}
                >
                  {completedTasks}
                </Text>
              </Stack>
            </Stack.Item>

            <Stack.Item grow className={styles.metricCard}>
              <Stack tokens={{ childrenGap: 8 }}>
                <Text variant="large" className={styles.cardTitle}>
                  Pending
                </Text>
                <Text
                  variant="superLarge"
                  className={`${styles.cardValue} ${styles.warningColor}`}
                >
                  {pendingTasks}
                </Text>
              </Stack>
            </Stack.Item>
          </Stack>

          <Stack.Item className={styles.metricCard}>
            <Stack tokens={{ childrenGap: 12 }}>
              <Text variant="large">Recent Tasks</Text>
              {tasks.length === 0 ? (
                <Text>No tasks found.</Text>
              ) : (
                <Stack tokens={{ childrenGap: 8 }}>
                  {tasks.slice(0, 5).map((task) => (
                    <Stack
                      key={task.Id}
                      horizontal
                      horizontalAlign="space-between"
                      className={styles.taskRow}
                    >
                      <Stack.Item grow>
                        <Text>{task.Title}</Text>
                      </Stack.Item>
                      <Stack.Item>
                        <Text
                          variant="small"
                          className={
                            task.Status === 'Completed'
                              ? styles.statusCompleted
                              : styles.statusPending
                          }
                        >
                          {task.Status || 'Unknown'}
                        </Text>
                      </Stack.Item>
                    </Stack>
                  ))}
                </Stack>
              )}
            </Stack>
          </Stack.Item>
        </>
      )}
    </Stack>
  );
};

export default Dashboard;
