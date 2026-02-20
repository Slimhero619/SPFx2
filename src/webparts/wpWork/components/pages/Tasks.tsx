/**
 * Tasks.tsx
 * Tasks management page component
 * View, create, and manage tasks from VATaskList
 */

import * as React from 'react';
import {
  Stack,
  Text,
  DetailsList,
  IColumn,
  SelectionMode,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Dialog,
  DialogType,
  TextField,
  DialogFooter,
  DefaultButton,
  Dropdown,
  IDropdownOption,
} from '@fluentui/react';
import styles from '../AppBoilerplate.module.scss';
import { SPService, ITaskItem } from '../../services/SPService';
import { IPageComponentProps } from './IPageComponentProps';

export interface ITasksProps extends IPageComponentProps {}

const columns: IColumn[] = [
  {
    key: 'Title',
    name: 'Title',
    fieldName: 'Title',
    minWidth: 200,
    isResizable: true,
  },
  {
    key: 'Status',
    name: 'Status',
    fieldName: 'Status',
    minWidth: 100,
    isResizable: true,
  },
  {
    key: 'DueDate',
    name: 'Due Date',
    fieldName: 'DueDate',
    minWidth: 120,
    isResizable: true,
  },
];

const statusOptions: IDropdownOption[] = [
  { key: 'Pending', text: 'Pending' },
  { key: 'In Progress', text: 'In Progress' },
  { key: 'Completed', text: 'Completed' },
];

export const Tasks: React.FC<ITasksProps> = ({ spService }) => {
  const [tasks, setTasks] = React.useState<ITaskItem[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [showNewTaskDialog, setShowNewTaskDialog] = React.useState(false);
  const [newTaskTitle, setNewTaskTitle] = React.useState('');
  const [newTaskStatus, setNewTaskStatus] = React.useState('Pending');
  const [newTaskDueDate, setNewTaskDueDate] = React.useState('');

  React.useEffect(() => {
    fetchTasks();
  }, [spService]);

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

  const handleAddTask = async () => {
    if (!newTaskTitle.trim()) {
      setError('Task title is required.');
      return;
    }

    try {
      await spService.createTask({
        Title: newTaskTitle,
        Status: newTaskStatus,
        DueDate: newTaskDueDate || undefined,
      });
      setShowNewTaskDialog(false);
      setNewTaskTitle('');
      setNewTaskStatus('Pending');
      setNewTaskDueDate('');
      fetchTasks();
    } catch (err) {
      setError('Failed to create task. Please try again.');
      console.error(err);
    }
  };

  return (
    <Stack className={styles.pageContent} tokens={{ childrenGap: 20 }}>
      <Stack.Item>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Stack.Item>
            <Text variant="xLarge" block>
              Tasks Management
            </Text>
            <Text variant="small" className={styles.subtitle}>
              Manage all tasks from VATaskList
            </Text>
          </Stack.Item>
          <Stack.Item>
            <PrimaryButton
              text="+ New Task"
              onClick={() => setShowNewTaskDialog(true)}
              disabled={loading}
            />
          </Stack.Item>
        </Stack>
      </Stack.Item>

      {error && (
        <Stack.Item>
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        </Stack.Item>
      )}

      {loading ? (
        <Stack.Item>
          <MessageBar messageBarType={MessageBarType.info}>
            Loading tasks...
          </MessageBar>
        </Stack.Item>
      ) : (
        <Stack.Item className={styles.metricCard}>
          {tasks.length === 0 ? (
            <Text>No tasks found. Click the "+ New Task" button to create one.</Text>
          ) : (
            <DetailsList
              items={tasks}
              columns={columns}
              selectionMode={SelectionMode.none}
              getKey={(item) => item.Id}
            />
          )}
        </Stack.Item>
      )}

      <Dialog
        hidden={!showNewTaskDialog}
        onDismiss={() => setShowNewTaskDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Create New Task',
          closeButtonAriaLabel: 'Close',
        }}
      >
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField
            label="Task Title"
            value={newTaskTitle}
            onChange={(e, value) => setNewTaskTitle(value || '')}
            required
          />
          <Dropdown
            label="Status"
            options={statusOptions}
            selectedKey={newTaskStatus}
            onChange={(e, option) => setNewTaskStatus(option?.key as string)}
          />
          <TextField
            label="Due Date"
            type="date"
            value={newTaskDueDate}
            onChange={(e, value) => setNewTaskDueDate(value || '')}
          />
        </Stack>
        <DialogFooter>
          <PrimaryButton onClick={handleAddTask} text="Create" />
          <DefaultButton onClick={() => setShowNewTaskDialog(false)} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

export default Tasks;
