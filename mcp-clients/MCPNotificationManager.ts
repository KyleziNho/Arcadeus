/**
 * MCPNotificationManager.ts
 * Manages MCP protocol notifications between servers and clients
 */

export interface MCPNotification {
  id: string;
  method: string;
  params?: any;
  timestamp: Date;
  serverId: string;
  priority: 'low' | 'normal' | 'high' | 'critical';
  category: string;
  title: string;
  message: string;
  data?: any;
  actions?: NotificationAction[];
  persistent?: boolean;
  autoHide?: number; // ms
  read?: boolean;
}

export interface NotificationAction {
  id: string;
  label: string;
  type: 'button' | 'link' | 'dismiss';
  action: string;
  data?: any;
}

export interface NotificationSubscription {
  id: string;
  serverId?: string;
  methods: string[];
  categories: string[];
  priorities: string[];
  callback: (notification: MCPNotification) => Promise<void>;
  filter?: (notification: MCPNotification) => boolean;
}

export interface NotificationUIConfig {
  position: 'top-right' | 'top-left' | 'bottom-right' | 'bottom-left';
  maxVisible: number;
  defaultAutoHide: number;
  showIcons: boolean;
  enableSound: boolean;
  groupSimilar: boolean;
}

export class MCPNotificationManager {
  private notifications: MCPNotification[] = [];
  private subscriptions: Map<string, NotificationSubscription> = new Map();
  private uiConfig: NotificationUIConfig;
  private maxStoredNotifications: number = 500;
  private notificationContainer: HTMLElement | null = null;
  private visibleNotifications: Map<string, HTMLElement> = new Map();

  constructor(config?: Partial<NotificationUIConfig>) {
    this.uiConfig = {
      position: 'top-right',
      maxVisible: 5,
      defaultAutoHide: 5000,
      showIcons: true,
      enableSound: false,
      groupSimilar: true,
      ...config
    };

    this.initializeUI();
  }

  /**
   * Initialize notification UI
   */
  private initializeUI(): void {
    this.notificationContainer = document.createElement('div');
    this.notificationContainer.className = `mcp-notifications mcp-notifications-${this.uiConfig.position}`;
    this.notificationContainer.innerHTML = `
      <style>
        .mcp-notifications {
          position: fixed;
          z-index: 10000;
          pointer-events: none;
          max-width: 400px;
        }
        
        .mcp-notifications-top-right {
          top: 20px;
          right: 20px;
        }
        
        .mcp-notifications-top-left {
          top: 20px;
          left: 20px;
        }
        
        .mcp-notifications-bottom-right {
          bottom: 20px;
          right: 20px;
        }
        
        .mcp-notifications-bottom-left {
          bottom: 20px;
          left: 20px;
        }
        
        .mcp-notification {
          pointer-events: auto;
          background: white;
          border-radius: 8px;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
          margin-bottom: 12px;
          overflow: hidden;
          transform: translateX(400px);
          opacity: 0;
          transition: all 0.3s ease;
        }
        
        .mcp-notification.show {
          transform: translateX(0);
          opacity: 1;
        }
        
        .mcp-notification.hide {
          transform: translateX(400px);
          opacity: 0;
        }
        
        .mcp-notification-header {
          display: flex;
          align-items: center;
          padding: 12px 16px 8px;
          border-left: 4px solid #007acc;
        }
        
        .mcp-notification.priority-low .mcp-notification-header {
          border-left-color: #28a745;
        }
        
        .mcp-notification.priority-normal .mcp-notification-header {
          border-left-color: #007acc;
        }
        
        .mcp-notification.priority-high .mcp-notification-header {
          border-left-color: #ffc107;
        }
        
        .mcp-notification.priority-critical .mcp-notification-header {
          border-left-color: #dc3545;
        }
        
        .mcp-notification-icon {
          width: 20px;
          height: 20px;
          margin-right: 8px;
          flex-shrink: 0;
        }
        
        .mcp-notification-title {
          font-weight: 600;
          font-size: 14px;
          color: #2c3e50;
          flex-grow: 1;
        }
        
        .mcp-notification-close {
          background: none;
          border: none;
          font-size: 18px;
          cursor: pointer;
          color: #666;
          padding: 0;
          width: 20px;
          height: 20px;
        }
        
        .mcp-notification-close:hover {
          color: #333;
        }
        
        .mcp-notification-body {
          padding: 0 16px 12px;
          font-size: 13px;
          color: #555;
          line-height: 1.4;
        }
        
        .mcp-notification-actions {
          display: flex;
          gap: 8px;
          padding: 8px 16px 12px;
          border-top: 1px solid #f0f0f0;
        }
        
        .mcp-notification-action {
          padding: 6px 12px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 12px;
          transition: background-color 0.2s;
        }
        
        .mcp-notification-action.primary {
          background: #007acc;
          color: white;
        }
        
        .mcp-notification-action.primary:hover {
          background: #005a9e;
        }
        
        .mcp-notification-action.secondary {
          background: #f8f9fa;
          color: #495057;
          border: 1px solid #dee2e6;
        }
        
        .mcp-notification-action.secondary:hover {
          background: #e9ecef;
        }
        
        .mcp-notification-time {
          font-size: 11px;
          color: #999;
          margin-top: 4px;
        }
      </style>
    `;

    document.body.appendChild(this.notificationContainer);
  }

  /**
   * Handle incoming MCP notification
   */
  async handleNotification(
    serverId: string,
    method: string,
    params: any
  ): Promise<void> {
    const notification: MCPNotification = {
      id: this.generateNotificationId(),
      method,
      params,
      timestamp: new Date(),
      serverId,
      priority: this.determinePriority(method, params),
      category: this.determineCategory(method),
      title: this.generateTitle(method, params),
      message: this.generateMessage(method, params),
      data: params,
      persistent: this.isPersistent(method),
      autoHide: this.getAutoHideDelay(method),
      read: false
    };

    // Add actions based on notification type
    notification.actions = this.generateActions(notification);

    // Store notification
    this.addNotification(notification);

    // Trigger subscriptions
    await this.triggerSubscriptions(notification);

    // Show in UI
    this.showNotificationUI(notification);

    console.log(`🔔 MCP Notification: ${notification.title}`);
  }

  /**
   * Subscribe to notifications
   */
  subscribe(
    options: Partial<NotificationSubscription> & {
      callback: (notification: MCPNotification) => Promise<void>;
    }
  ): string {
    const subscriptionId = this.generateSubscriptionId();
    
    const subscription: NotificationSubscription = {
      id: subscriptionId,
      methods: options.methods || ['*'],
      categories: options.categories || ['*'],
      priorities: options.priorities || ['*'],
      callback: options.callback,
      filter: options.filter,
      serverId: options.serverId
    };

    this.subscriptions.set(subscriptionId, subscription);
    
    console.log(`📧 Subscribed to notifications: ${subscriptionId}`);
    return subscriptionId;
  }

  /**
   * Unsubscribe from notifications
   */
  unsubscribe(subscriptionId: string): boolean {
    return this.subscriptions.delete(subscriptionId);
  }

  /**
   * Send notification (for servers to notify clients)
   */
  async sendNotification(notification: Partial<MCPNotification>): Promise<void> {
    const fullNotification: MCPNotification = {
      id: notification.id || this.generateNotificationId(),
      method: notification.method || 'notification',
      timestamp: new Date(),
      serverId: notification.serverId || 'local',
      priority: notification.priority || 'normal',
      category: notification.category || 'general',
      title: notification.title || 'Notification',
      message: notification.message || '',
      data: notification.data,
      actions: notification.actions,
      persistent: notification.persistent,
      autoHide: notification.autoHide,
      read: false,
      ...notification
    };

    await this.handleNotification(
      fullNotification.serverId,
      fullNotification.method,
      fullNotification.data
    );
  }

  /**
   * Get all notifications
   */
  getNotifications(filter?: {
    unreadOnly?: boolean;
    serverId?: string;
    category?: string;
    priority?: string;
    since?: Date;
  }): MCPNotification[] {
    let notifications = [...this.notifications];

    if (filter) {
      if (filter.unreadOnly) {
        notifications = notifications.filter(n => !n.read);
      }
      if (filter.serverId) {
        notifications = notifications.filter(n => n.serverId === filter.serverId);
      }
      if (filter.category) {
        notifications = notifications.filter(n => n.category === filter.category);
      }
      if (filter.priority) {
        notifications = notifications.filter(n => n.priority === filter.priority);
      }
      if (filter.since) {
        notifications = notifications.filter(n => n.timestamp >= filter.since!);
      }
    }

    return notifications.sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime());
  }

  /**
   * Mark notification as read
   */
  markAsRead(notificationId: string): boolean {
    const notification = this.notifications.find(n => n.id === notificationId);
    if (notification) {
      notification.read = true;
      this.updateVisibleNotification(notificationId);
      return true;
    }
    return false;
  }

  /**
   * Mark all notifications as read
   */
  markAllAsRead(serverId?: string): void {
    this.notifications
      .filter(n => !serverId || n.serverId === serverId)
      .forEach(n => n.read = true);
    
    // Update all visible notifications
    for (const id of this.visibleNotifications.keys()) {
      this.updateVisibleNotification(id);
    }
  }

  /**
   * Dismiss notification
   */
  dismissNotification(notificationId: string): void {
    const element = this.visibleNotifications.get(notificationId);
    if (element) {
      element.classList.add('hide');
      setTimeout(() => {
        element.remove();
        this.visibleNotifications.delete(notificationId);
      }, 300);
    }
  }

  /**
   * Clear old notifications
   */
  clearNotifications(olderThan?: Date): void {
    const cutoff = olderThan || new Date(Date.now() - 7 * 24 * 60 * 60 * 1000); // 7 days
    
    this.notifications = this.notifications.filter(n => {
      if (n.timestamp < cutoff && !n.persistent) {
        // Also remove from visible notifications
        this.dismissNotification(n.id);
        return false;
      }
      return true;
    });

    console.log(`🗑️ Cleared old notifications before ${cutoff.toISOString()}`);
  }

  /**
   * Get notification statistics
   */
  getStatistics(): {
    total: number;
    unread: number;
    byPriority: { [key: string]: number };
    byCategory: { [key: string]: number };
    byServer: { [key: string]: number };
  } {
    const stats = {
      total: this.notifications.length,
      unread: this.notifications.filter(n => !n.read).length,
      byPriority: {} as { [key: string]: number },
      byCategory: {} as { [key: string]: number },
      byServer: {} as { [key: string]: number }
    };

    this.notifications.forEach(n => {
      stats.byPriority[n.priority] = (stats.byPriority[n.priority] || 0) + 1;
      stats.byCategory[n.category] = (stats.byCategory[n.category] || 0) + 1;
      stats.byServer[n.serverId] = (stats.byServer[n.serverId] || 0) + 1;
    });

    return stats;
  }

  // Private helper methods

  private addNotification(notification: MCPNotification): void {
    this.notifications.push(notification);
    
    // Maintain storage limit
    if (this.notifications.length > this.maxStoredNotifications) {
      this.notifications.splice(0, this.notifications.length - this.maxStoredNotifications);
    }
  }

  private async triggerSubscriptions(notification: MCPNotification): Promise<void> {
    for (const subscription of this.subscriptions.values()) {
      if (this.matchesSubscription(notification, subscription)) {
        try {
          await subscription.callback(notification);
        } catch (error) {
          console.error('Notification subscription callback failed:', error);
        }
      }
    }
  }

  private matchesSubscription(
    notification: MCPNotification,
    subscription: NotificationSubscription
  ): boolean {
    // Server filter
    if (subscription.serverId && subscription.serverId !== notification.serverId) {
      return false;
    }

    // Method filter
    if (!subscription.methods.includes('*') && 
        !subscription.methods.includes(notification.method)) {
      return false;
    }

    // Category filter
    if (!subscription.categories.includes('*') && 
        !subscription.categories.includes(notification.category)) {
      return false;
    }

    // Priority filter
    if (!subscription.priorities.includes('*') && 
        !subscription.priorities.includes(notification.priority)) {
      return false;
    }

    // Custom filter
    if (subscription.filter && !subscription.filter(notification)) {
      return false;
    }

    return true;
  }

  private showNotificationUI(notification: MCPNotification): void {
    if (this.visibleNotifications.size >= this.uiConfig.maxVisible) {
      // Remove oldest visible notification
      const oldestId = Array.from(this.visibleNotifications.keys())[0];
      this.dismissNotification(oldestId);
    }

    const element = this.createNotificationElement(notification);
    this.notificationContainer?.appendChild(element);
    this.visibleNotifications.set(notification.id, element);

    // Trigger show animation
    setTimeout(() => {
      element.classList.add('show');
    }, 50);

    // Auto-hide if configured
    if (notification.autoHide) {
      setTimeout(() => {
        this.dismissNotification(notification.id);
      }, notification.autoHide);
    }
  }

  private createNotificationElement(notification: MCPNotification): HTMLElement {
    const element = document.createElement('div');
    element.className = `mcp-notification priority-${notification.priority}`;
    element.setAttribute('data-id', notification.id);

    const icon = this.getIconForCategory(notification.category);
    const timeAgo = this.formatTimeAgo(notification.timestamp);

    element.innerHTML = `
      <div class="mcp-notification-header">
        ${this.uiConfig.showIcons ? `<div class="mcp-notification-icon">${icon}</div>` : ''}
        <div class="mcp-notification-title">${this.escapeHtml(notification.title)}</div>
        <button class="mcp-notification-close" onclick="this.closest('.mcp-notification').remove()">×</button>
      </div>
      <div class="mcp-notification-body">
        ${this.escapeHtml(notification.message)}
        <div class="mcp-notification-time">${timeAgo}</div>
      </div>
      ${notification.actions && notification.actions.length > 0 ? `
        <div class="mcp-notification-actions">
          ${notification.actions.map(action => `
            <button class="mcp-notification-action ${action.type === 'button' ? 'primary' : 'secondary'}"
                    onclick="window.mcpNotificationAction('${notification.id}', '${action.id}')">
              ${this.escapeHtml(action.label)}
            </button>
          `).join('')}
        </div>
      ` : ''}
    `;

    // Set up global action handler
    (window as any).mcpNotificationAction = (notificationId: string, actionId: string) => {
      this.handleNotificationAction(notificationId, actionId);
    };

    return element;
  }

  private updateVisibleNotification(notificationId: string): void {
    const element = this.visibleNotifications.get(notificationId);
    if (element) {
      const notification = this.notifications.find(n => n.id === notificationId);
      if (notification?.read) {
        element.style.opacity = '0.7';
      }
    }
  }

  private handleNotificationAction(notificationId: string, actionId: string): void {
    const notification = this.notifications.find(n => n.id === notificationId);
    const action = notification?.actions?.find(a => a.id === actionId);
    
    if (notification && action) {
      console.log(`🎬 Notification action: ${action.action}`);
      
      // Mark as read
      this.markAsRead(notificationId);
      
      // Execute action
      switch (action.action) {
        case 'dismiss':
          this.dismissNotification(notificationId);
          break;
        case 'open-chat':
          // Would open chat interface
          break;
        case 'show-details':
          // Would show more details
          break;
        default:
          console.log('Unknown notification action:', action.action);
      }
    }
  }

  private determinePriority(method: string, params: any): MCPNotification['priority'] {
    if (method.includes('error') || method.includes('failed')) return 'critical';
    if (method.includes('warning') || method.includes('timeout')) return 'high';
    if (method.includes('progress') || method.includes('status')) return 'low';
    return 'normal';
  }

  private determineCategory(method: string): string {
    if (method.includes('tools')) return 'tools';
    if (method.includes('resources')) return 'resources';
    if (method.includes('error')) return 'error';
    if (method.includes('progress')) return 'progress';
    if (method.includes('excel')) return 'excel';
    return 'general';
  }

  private generateTitle(method: string, params: any): string {
    const titles: { [key: string]: string } = {
      'tools/list_changed': 'Tools Updated',
      'resources/list_changed': 'Resources Updated',
      'progress/updated': 'Operation Progress',
      'error/occurred': 'Error Occurred',
      'excel/data_changed': 'Excel Data Changed'
    };
    
    return titles[method] || 'MCP Notification';
  }

  private generateMessage(method: string, params: any): string {
    const messages: { [key: string]: string } = {
      'tools/list_changed': 'Available tools have been updated',
      'resources/list_changed': 'Available resources have changed',
      'progress/updated': `Progress: ${params?.percentage || 0}%`,
      'error/occurred': params?.message || 'An error occurred',
      'excel/data_changed': `Data changed in ${params?.worksheet || 'worksheet'}`
    };
    
    return messages[method] || params?.message || 'MCP notification received';
  }

  private isPersistent(method: string): boolean {
    return method.includes('error') || method.includes('critical');
  }

  private getAutoHideDelay(method: string): number {
    if (method.includes('error')) return 10000; // 10 seconds
    if (method.includes('progress')) return 3000; // 3 seconds
    return this.uiConfig.defaultAutoHide;
  }

  private generateActions(notification: MCPNotification): NotificationAction[] {
    const actions: NotificationAction[] = [];
    
    if (notification.category === 'error') {
      actions.push({
        id: 'show-details',
        label: 'Details',
        type: 'button',
        action: 'show-details'
      });
    }
    
    actions.push({
      id: 'dismiss',
      label: 'Dismiss',
      type: 'button',
      action: 'dismiss'
    });
    
    return actions;
  }

  private getIconForCategory(category: string): string {
    const icons: { [key: string]: string } = {
      'error': '❌',
      'excel': '📊',
      'tools': '🔧',
      'resources': '📁',
      'progress': '⏳',
      'general': '🔔'
    };
    
    return icons[category] || '🔔';
  }

  private formatTimeAgo(date: Date): string {
    const seconds = Math.floor((Date.now() - date.getTime()) / 1000);
    
    if (seconds < 60) return 'just now';
    if (seconds < 3600) return `${Math.floor(seconds / 60)}m ago`;
    if (seconds < 86400) return `${Math.floor(seconds / 3600)}h ago`;
    return `${Math.floor(seconds / 86400)}d ago`;
  }

  private escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  private generateNotificationId(): string {
    return `notif_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  private generateSubscriptionId(): string {
    return `sub_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }
}

export default MCPNotificationManager;