/**
 * ============================================================================
 * 15_EventBus.gs - EVENT BUS / PUB-SUB SYSTEM
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Decoupled publish/subscribe event bus for inter-module communication.
 *   Modules subscribe to named events (e.g., 'sheet:edit:GRIEVANCE_LOG')
 *   and react independently. Supports wildcard listeners that receive all
 *   events. Maintains an event log (last 200 events) for debugging. Can be
 *   enabled/disabled globally.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Replaces direct function calls in trigger handlers, which caused tight
 *   coupling between modules. With the EventBus, adding a new reaction to
 *   a sheet edit doesn't require modifying onEdit() — just subscribe from
 *   your module. Event naming uses domain:action convention for clarity.
 *   The wildcard listener feature allows audit logging to capture all events
 *   without subscribing individually.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   onEdit() dispatched events won't reach subscriber modules. Real-time
 *   features (auto-formatting, status tracking, badge refresh) stop
 *   updating. Core CRUD operations still work because they don't depend
 *   on the EventBus — it's used for secondary/reactive behavior only.
 *
 * DEPENDENCIES:
 *   Depends on nothing (self-contained IIFE).
 *   Used by 10_Main.gs (onEdit dispatches events), and any module that
 *   subscribes to sheet edit events.
 *
 * @version 4.43.1
 * @license Free for use by non-profit collective bargaining groups and unions
 * ============================================================================
 */

// ============================================================================
// EVENT BUS CORE
// ============================================================================

/**
 * Central event bus for decoupled module communication.
 * Supports subscribe, unsubscribe, emit, and wildcard listeners.
 */
var EventBus = (function() {
  var listeners_ = {};
  var wildcardListeners_ = [];
  var eventLog_ = [];
  var MAX_LOG_SIZE = 200;
  var enabled_ = true;

  return {
    /**
     * Subscribe to a named event
     * @param {string} eventName - Event to listen for (e.g. 'sheet:edit:GRIEVANCE_LOG')
     * @param {Function} callback - Handler function receiving event data
     * @param {Object} [options] - Optional: { priority: number, once: boolean, id: string }
     * @returns {string} Subscription ID for later removal
     */
    on: function(eventName, callback, options) {
      options = options || {};
      if (!listeners_[eventName]) {
        listeners_[eventName] = [];
      }

      // Collision probability: ~1 in 60M per subscription pair (5 random chars from base-36
      // = 36^5 = ~60M combinations, plus millisecond timestamp). Acceptable for current
      // scale (~20 active subscribers per session).
      var subId = options.id || ('sub_' + eventName + '_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7));

      listeners_[eventName].push({
        id: subId,
        callback: callback,
        priority: options.priority || 0,
        once: options.once || false
      });

      // Sort by priority (higher runs first)
      listeners_[eventName].sort(function(a, b) { return b.priority - a.priority; });

      return subId;
    },

    /**
     * Subscribe to a named event, auto-removing after first invocation
     * @param {string} eventName - Event to listen for
     * @param {Function} callback - Handler function
     * @returns {string} Subscription ID
     */
    once: function(eventName, callback) {
      return this.on(eventName, callback, { once: true });
    },

    /**
     * Subscribe to ALL events (wildcard listener)
     * Useful for audit logging and debugging
     * @param {Function} callback - Receives (eventName, data)
     * @returns {string} Subscription ID
     */
    onAny: function(callback) {
      var subId = 'wildcard_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7);
      wildcardListeners_.push({ id: subId, callback: callback });
      return subId;
    },

    /**
     * Remove a subscription by ID
     * @param {string} subId - The subscription ID returned by on()
     */
    off: function(subId) {
      Object.keys(listeners_).forEach(function(eventName) {
        listeners_[eventName] = listeners_[eventName].filter(function(sub) {
          return sub.id !== subId;
        });
      });
      wildcardListeners_ = wildcardListeners_.filter(function(sub) {
        return sub.id !== subId;
      });
    },

    /**
     * Remove all listeners for a given event
     * @param {string} eventName - Event name to clear
     */
    offAll: function(eventName) {
      if (eventName) {
        delete listeners_[eventName];
      } else {
        listeners_ = {};
        wildcardListeners_ = [];
      }
    },

    /**
     * Emit an event, invoking all registered listeners
     * @param {string} eventName - Event to emit
     * @param {*} [data] - Data payload passed to each listener
     * @returns {Object} Result summary: { handled: number, errors: string[] }
     */
    emit: function(eventName, data) {
      if (!enabled_) return { handled: 0, errors: [] };

      var result = { handled: 0, errors: [] };
      var timestamp = new Date().toISOString();

      // Log the event
      eventLog_.push({ event: eventName, timestamp: timestamp, data: data ? '(payload)' : null });
      if (eventLog_.length > MAX_LOG_SIZE) {
        eventLog_ = eventLog_.slice(-MAX_LOG_SIZE);
      }

      // Invoke wildcard listeners
      wildcardListeners_.forEach(function(sub) {
        try {
          sub.callback(eventName, data);
          result.handled++;
        } catch (err) {
          result.errors.push('Wildcard[' + sub.id + ']: ' + err.message);
          Logger.log('EventBus wildcard error: ' + err.message);
        }
      });

      // Invoke specific listeners
      var subs = listeners_[eventName] || [];
      var toRemove = [];

      for (var i = 0; i < subs.length; i++) {
        try {
          subs[i].callback(data);
          result.handled++;
        } catch (err) {
          result.errors.push(eventName + '[' + subs[i].id + ']: ' + err.message);
          Logger.log('EventBus error in ' + eventName + ': ' + err.message);
        }
        if (subs[i].once) {
          toRemove.push(subs[i].id);
        }
      }

      // Also match domain-level wildcards: 'sheet:edit' matches 'sheet:edit:GRIEVANCE_LOG'
      var parts = eventName.split(':');
      if (parts.length > 2) {
        var parentEvent = parts.slice(0, 2).join(':');
        var parentSubs = listeners_[parentEvent] || [];
        for (var j = 0; j < parentSubs.length; j++) {
          try {
            parentSubs[j].callback(data);
            result.handled++;
          } catch (err) {
            result.errors.push(parentEvent + '[' + parentSubs[j].id + ']: ' + err.message);
            Logger.log('EventBus error in ' + parentEvent + ' (parent of ' + eventName + '): ' + err.message);
          }
          if (parentSubs[j].once) {
            toRemove.push(parentSubs[j].id);
          }
        }
      }

      // Clean up once-only subscriptions
      for (var k = 0; k < toRemove.length; k++) {
        this.off(toRemove[k]);
      }

      return result;
    },

    /**
     * Enable or disable the event bus globally
     * @param {boolean} state
     */
    setEnabled: function(state) {
      enabled_ = !!state;
    },

    /**
     * Get the event log (recent events)
     * @param {number} [count] - Number of recent events (default: all)
     * @returns {Array<Object>} Event log entries
     */
    getLog: function(count) {
      if (count) {
        return eventLog_.slice(-count);
      }
      return eventLog_.slice();
    },

    /**
     * Get listener count for an event (or total)
     * @param {string} [eventName] - Specific event, or omit for total
     * @returns {number} Number of active listeners
     */
    listenerCount: function(eventName) {
      if (eventName) {
        return (listeners_[eventName] || []).length;
      }
      var total = wildcardListeners_.length;
      Object.keys(listeners_).forEach(function(key) {
        total += listeners_[key].length;
      });
      return total;
    },

    /**
     * List all registered event names
     * @returns {string[]} Array of event names with active listeners
     */
    eventNames: function() {
      return Object.keys(listeners_).filter(function(key) {
        return listeners_[key].length > 0;
      });
    },

    /**
     * Reset the event bus (clear all listeners and log)
     */
    reset: function() {
      listeners_ = {};
      wildcardListeners_ = [];
      eventLog_ = [];
      enabled_ = true;
    }
  };
})();

// ============================================================================
// EVENT REGISTRATION — Module subscribers
// ============================================================================

/**
 * Register all default event subscriptions.
 * Called once during initialization (onOpen or trigger setup).
 * Each module registers its own listeners here.
 */
var eventBusSubscribersRegistered_ = false;

function registerEventBusSubscribers() {
  // Only reset if this is the first registration to avoid wiping prior subscribers
  if (!eventBusSubscribersRegistered_) {
    EventBus.reset();
    eventBusSubscribersRegistered_ = true;
  } else {
    // Already registered — skip to avoid duplicate listeners
    return;
  }

  // --- Grievance Log edit handlers ---
  EventBus.on('sheet:edit:GRIEVANCE_LOG', function(e) {
    if (typeof handleGrievanceEdit === 'function') handleGrievanceEdit(e);
  }, { priority: 100, id: 'grievance_edit_handler' });

  EventBus.on('sheet:edit:GRIEVANCE_LOG', function(e) {
    if (typeof applyAutoStyleToRow_ === 'function') applyAutoStyleToRow_(e.range.getSheet(), e.range.getRow());
  }, { priority: 90, id: 'grievance_style_handler' });

  EventBus.on('sheet:edit:GRIEVANCE_LOG', function(e) {
    if (typeof handleStageGateWorkflow_ === 'function') handleStageGateWorkflow_(e);
  }, { priority: 80, id: 'grievance_stagegate_handler' });

  EventBus.on('sheet:edit:GRIEVANCE_LOG', function(e) {
    try { syncDropdownToConfig_(e, e.range.getSheet().getName()); } catch (_err) { Logger.log('_err: ' + (_err.message || _err)); }
  }, { priority: 70, id: 'grievance_config_sync' });

  EventBus.on('sheet:edit:GRIEVANCE_LOG', function() {
    if (typeof sortGrievanceLogByStatus === 'function') {
      try { sortGrievanceLogByStatus(); } catch (_err) { Logger.log('_err: ' + (_err.message || _err)); }
    }
  }, { priority: 60, id: 'grievance_sort_handler' });

  // --- Member Directory edit handlers ---
  EventBus.on('sheet:edit:MEMBER_DIR', function(e) {
    if (typeof handleMemberEdit === 'function') handleMemberEdit(e);
  }, { priority: 100, id: 'member_edit_handler' });

  EventBus.on('sheet:edit:MEMBER_DIR', function(e) {
    if (typeof applyAutoStyleToRow_ === 'function') applyAutoStyleToRow_(e.range.getSheet(), e.range.getRow());
  }, { priority: 90, id: 'member_style_handler' });

  EventBus.on('sheet:edit:MEMBER_DIR', function(e) {
    try { syncDropdownToConfig_(e, e.range.getSheet().getName()); } catch (_err) { Logger.log('_err: ' + (_err.message || _err)); }
  }, { priority: 70, id: 'member_config_sync' });

  // --- Checklist edit handler ---
  EventBus.on('sheet:edit:CASE_CHECKLIST', function(e) {
    if (typeof handleChecklistEdit === 'function') handleChecklistEdit(e);
  }, { priority: 100, id: 'checklist_edit_handler' });

  // --- Volunteer Hours sync ---
  EventBus.on('sheet:edit:VOLUNTEER_HOURS', function() {
    if (typeof syncVolunteerHoursToMemberDirectory === 'function') {
      try { syncVolunteerHoursToMemberDirectory(); } catch (_err) { Logger.log('_err: ' + (_err.message || _err)); }
    }
  }, { priority: 100, id: 'volunteer_sync_handler' });

  // --- Meeting Attendance sync ---
  EventBus.on('sheet:edit:MEETING_ATTENDANCE', function() {
    if (typeof syncMeetingAttendanceToMemberDirectory === 'function') {
      try { syncMeetingAttendanceToMemberDirectory(); } catch (_err) { Logger.log('_err: ' + (_err.message || _err)); }
    }
  }, { priority: 100, id: 'attendance_sync_handler' });

  // --- Config edit handlers ---
  EventBus.on('sheet:edit:CONFIG', function(e) {
    if (typeof handleConfigStewardEdit_ === 'function') handleConfigStewardEdit_(e);
  }, { priority: 100, id: 'config_steward_handler' });

  EventBus.on('sheet:edit:CONFIG', function(e) {
    if (typeof warnInvalidConfigValue_ === 'function') {
      try { warnInvalidConfigValue_(e); } catch (_err) { Logger.log('warnInvalidConfigValue_: ' + (_err.message || _err)); }
    }
  }, { priority: 90, id: 'config_type_validator' });

  EventBus.on('sheet:edit:CONFIG', function(e) {
    if (typeof syncConfigToSheetValidation_ === 'function') {
      try { syncConfigToSheetValidation_(e); } catch (_err) { Logger.log('syncConfigToSheetValidation_: ' + (_err.message || _err)); }
    }
  }, { priority: 80, id: 'config_validation_sync' });

  // --- Cross-cutting: Audit logging (high-value sheets) ---
  EventBus.on('sheet:edit:GRIEVANCE_LOG', function(e) {
    if (typeof onEditAudit === 'function') {
      try { onEditAudit(e); } catch (_err) { Logger.log('_err: ' + (_err.message || _err)); }
    }
  }, { priority: 10, id: 'grievance_audit' });

  EventBus.on('sheet:edit:MEMBER_DIR', function(e) {
    if (typeof onEditAudit === 'function') {
      try { onEditAudit(e); } catch (_err) { Logger.log('_err: ' + (_err.message || _err)); }
    }
  }, { priority: 10, id: 'member_audit' });

  // --- v4.30.0: Log terminal grievance status changes (archive candidates) ---
  EventBus.on('sheet:edit:GRIEVANCE_LOG', function(e) {
    try {
      if (!e || !e.range) return;
      var col = e.range.getColumn();
      if (typeof GRIEVANCE_COLS !== 'undefined' && col === GRIEVANCE_COLS.STATUS) {
        var newVal = String(e.range.getValue() || '').trim();
        var terminalStatuses = ['won', 'denied', 'settled', 'withdrawn', 'closed'];
        if (terminalStatuses.indexOf(newVal.toLowerCase()) !== -1) {
          var row = e.range.getRow();
          Logger.log('Grievance terminal status: row ' + row + ' → ' + newVal + ' (archive candidate)');
        }
      }
    } catch (_err) { Logger.log('_err: ' + (_err.message || _err)); }
  }, { priority: 50, id: 'grievance_terminal_status_logger' });

  // --- Auto-sync (debounced via onEditAutoSync) ---
  // M-48: Mark _grievanceEditHandled so onEditAutoSync skips redundant formula sync
  // (handleGrievanceEdit at priority 100 already computed deadline values for this row)
  EventBus.on('sheet:edit:GRIEVANCE_LOG', function(e) {
    if (typeof onEditAutoSync === 'function') {
      try { e._grievanceEditHandled = true; onEditAutoSync(e); } catch (_err) { Logger.log('_err: ' + (_err.message || _err)); }
    }
  }, { priority: 20, id: 'grievance_auto_sync' });

  // --- Data change notifications (for dashboard refresh) ---
  EventBus.on('data:changed', function(data) {
    Logger.log('Data changed: ' + (data && data.source ? data.source : 'unknown'));
  }, { priority: 0, id: 'data_change_logger' });

  // --- Auto-notification: Grievance deadline approaching ---
  EventBus.on('grievance:deadline:approaching', function(data) {
    try {
      if (typeof sendWebAppNotification === 'function') {
        sendWebAppNotification({
          recipient: data.memberEmail || 'All Stewards',
          type: 'Deadline',
          title: 'Grievance Deadline Approaching',
          message: data.grievanceId + ' (' + (data.memberName || 'Unknown') + ') \u2014 ' + data.daysLeft + ' day(s) remaining',
          priority: data.daysLeft <= 1 ? 'Urgent' : 'Normal'
        });
      }
    } catch (err) {
      Logger.log('EventBus notification error (deadline): ' + err);
    }
  }, { priority: 30, id: 'notif_deadline_approaching' });

  // --- Auto-notification: Grievance status changed ---
  EventBus.on('grievance:status:changed', function(data) {
    try {
      if (data.memberEmail && typeof sendWebAppNotification === 'function') {
        sendWebAppNotification({
          recipient: data.memberEmail,
          type: 'System',
          title: 'Grievance Update: ' + (data.grievanceId || ''),
          message: 'Status changed to ' + (data.newStatus || 'Unknown'),
          priority: 'Normal'
        });
      }
    } catch (err) {
      Logger.log('EventBus notification error (status): ' + err);
    }
  }, { priority: 30, id: 'notif_status_changed' });

  Logger.log('EventBus: ' + EventBus.listenerCount() + ' subscribers registered');
}

// ============================================================================
// EVENT BUS BRIDGE — Connects to existing onEdit dispatcher
// ============================================================================

/**
 * @private Builds and caches a reverse lookup map from sheet name to SHEETS constant key.
 * @returns {Object.<string, string>} Map of sheet names to their SHEETS constant keys
 */
// Module-level cache: rebuilt once per V8 execution context, not per edit event
var _emitSheetKeyMap = null;
function _getSheetKeyMap() {
  if (_emitSheetKeyMap) return _emitSheetKeyMap;
  _emitSheetKeyMap = {};
  Object.keys(SHEETS).forEach(function(key) {
    _emitSheetKeyMap[SHEETS[key]] = key;
  });
  return _emitSheetKeyMap;
}

/**
 * Routes a Sheets onEdit event through the EventBus to registered subscribers.
 * @param {Object} e - The Google Sheets edit event object
 * @returns {Object} Emit result with handled count and errors array
 */
function emitEditEvent(e) {
  if (!e || !e.range) return { handled: 0, errors: [] };

  var sheetName = e.range.getSheet().getName();
  var sheetKey = _getSheetKeyMap()[sheetName];
  if (!sheetKey) return { handled: 0, errors: [] };

  return EventBus.emit('sheet:edit:' + sheetKey, e);
}
// ============================================================================
// EVENT BUS DIAGNOSTICS
// ============================================================================

