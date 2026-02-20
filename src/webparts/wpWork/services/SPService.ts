/**
 * SPService.ts
 * PnP.js Service Layer for SharePoint data access
 * Initializes and manages spfi (SharePoint Fluent Interface) instance
 */

import { spfi } from '@pnp/sp';
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IWebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Interface for list item data from VATaskList
 */
export interface ITaskItem {
  Id: number;
  Title: string;
  Status?: string;
  DueDate?: string;
  AssignedTo?: string;
}

/**
 * SPService class - manages all SharePoint data operations
 */
export class SPService {
  private sp: SPFI;

  /**
   * Constructor initializes the spfi factory with SPFx context
   * @param context - SPFx web part context
   */
  constructor(context: IWebPartContext) {
    // Initialize spfi 
    // SPFx context automatically handles authentication
    this.sp = spfi();
  }

  /**
   * Fetch all tasks from VATaskList
   * @returns Promise of task items array
   */
  public async getTasksFromVATaskList(): Promise<ITaskItem[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle('VATaskList')
        .items
        .select('Id', 'Title', 'Status', 'DueDate')
        ();

      return items as ITaskItem[];
    } catch (error) {
      console.error('Error fetching tasks from VATaskList:', error);
      throw error;
    }
  }

  /**
   * Fetch a single task by ID
   * @param itemId - The ID of the task item
   * @returns Promise of a single task item
   */
  public async getTaskById(itemId: number): Promise<ITaskItem> {
    try {
      const item = await this.sp.web.lists
        .getByTitle('VATaskList')
        .items.getById(itemId)
        .select('Id', 'Title', 'Status', 'DueDate')
        ();

      return item as ITaskItem;
    } catch (error) {
      console.error(`Error fetching task with ID ${itemId}:`, error);
      throw error;
    }
  }

  /**
   * Create a new task in VATaskList
   * @param taskData - Task data object
   * @returns Promise of the created item
   */
  public async createTask(taskData: Partial<ITaskItem>): Promise<ITaskItem> {
    try {
      const result = await this.sp.web.lists
        .getByTitle('VATaskList')
        .items.add({
          Title: taskData.Title,
          Status: taskData.Status || 'Pending',
          DueDate: taskData.DueDate,
        });

      return { Id: result.data.Id, ...taskData } as ITaskItem;
    } catch (error) {
      console.error('Error creating task:', error);
      throw error;
    }
  }

  /**
   * Update an existing task in VATaskList
   * @param itemId - The ID of the task item to update
   * @param taskData - Updated task data
   * @returns Promise of update completion
   */
  public async updateTask(
    itemId: number,
    taskData: Partial<ITaskItem>
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('VATaskList')
        .items.getById(itemId)
        .update(taskData);
    } catch (error) {
      console.error(`Error updating task with ID ${itemId}:`, error);
      throw error;
    }
  }

  /**
   * Delete a task from VATaskList
   * @param itemId - The ID of the task item to delete
   * @returns Promise of deletion completion
   */
  public async deleteTask(itemId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('VATaskList')
        .items.getById(itemId)
        .delete();
    } catch (error) {
      console.error(`Error deleting task with ID ${itemId}:`, error);
      throw error;
    }
  }

  /**
   * Test method to verify SharePoint connectivity
   * @returns Promise of web title
   */
  public async testConnection(): Promise<string> {
    try {
      const web = await this.sp.web();
      return web.Title || 'Connected';
    } catch (error) {
      console.error('Error testing connection:', error);
      throw error;
    }
  }
}
