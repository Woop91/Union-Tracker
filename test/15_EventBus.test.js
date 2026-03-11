/**
 * Tests for 15_EventBus.gs
 *
 * Covers EventBus IIFE (subscribe, emit, wildcard, once, priority, logging),
 * event bridge functions, and subscriber registration.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '15_EventBus.gs']);

// Reset EventBus before each test
beforeEach(() => {
  EventBus.reset();
  global.eventBusSubscribersRegistered_ = false;
});

// ============================================================================
// EventBus Core
// ============================================================================

describe('EventBus.on / EventBus.emit', () => {
  test('registers and invokes a listener', () => {
    const handler = jest.fn();
    EventBus.on('test:event', handler);
    EventBus.emit('test:event', { data: 'hello' });
    expect(handler).toHaveBeenCalledWith({ data: 'hello' });
  });

  test('returns a subscription ID', () => {
    const subId = EventBus.on('test:event', jest.fn());
    expect(typeof subId).toBe('string');
    expect(subId.length).toBeGreaterThan(0);
  });

  test('supports custom subscription ID', () => {
    const subId = EventBus.on('test:event', jest.fn(), { id: 'custom_id' });
    expect(subId).toBe('custom_id');
  });

  test('emit returns handler count and errors', () => {
    EventBus.on('test:event', jest.fn());
    EventBus.on('test:event', jest.fn());
    const result = EventBus.emit('test:event');
    expect(result.handled).toBe(2);
    expect(result.errors).toEqual([]);
  });

  test('emit catches handler errors', () => {
    EventBus.on('test:event', () => { throw new Error('boom'); });
    const result = EventBus.emit('test:event');
    expect(result.errors.length).toBe(1);
    expect(result.errors[0]).toContain('boom');
  });

  test('emit returns {handled:0} for unregistered events', () => {
    const result = EventBus.emit('nonexistent');
    expect(result.handled).toBe(0);
  });
});

describe('EventBus.once', () => {
  test('fires listener only once', () => {
    const handler = jest.fn();
    EventBus.once('test:once', handler);
    EventBus.emit('test:once');
    EventBus.emit('test:once');
    expect(handler).toHaveBeenCalledTimes(1);
  });
});

describe('EventBus.off', () => {
  test('removes a specific listener', () => {
    const handler = jest.fn();
    const subId = EventBus.on('test:off', handler);
    EventBus.off(subId);
    EventBus.emit('test:off');
    expect(handler).not.toHaveBeenCalled();
  });
});

describe('EventBus.offAll', () => {
  test('removes all listeners for a specific event', () => {
    const handler1 = jest.fn();
    const handler2 = jest.fn();
    EventBus.on('test:offAll', handler1);
    EventBus.on('test:offAll', handler2);
    EventBus.offAll('test:offAll');
    EventBus.emit('test:offAll');
    expect(handler1).not.toHaveBeenCalled();
    expect(handler2).not.toHaveBeenCalled();
  });

  test('clears all listeners when no event specified', () => {
    EventBus.on('event1', jest.fn());
    EventBus.on('event2', jest.fn());
    EventBus.offAll();
    expect(EventBus.listenerCount()).toBe(0);
  });
});

describe('EventBus.onAny (wildcard)', () => {
  test('fires wildcard listener for any event', () => {
    const handler = jest.fn();
    EventBus.onAny(handler);
    EventBus.emit('any:event', 'payload');
    expect(handler).toHaveBeenCalledWith('any:event', 'payload');
  });

  test('wildcard listener receives all events', () => {
    const handler = jest.fn();
    EventBus.onAny(handler);
    EventBus.emit('first');
    EventBus.emit('second');
    expect(handler).toHaveBeenCalledTimes(2);
  });
});

describe('EventBus priority', () => {
  test('higher priority listeners fire first', () => {
    const order = [];
    EventBus.on('test:priority', () => order.push('low'), { priority: 10 });
    EventBus.on('test:priority', () => order.push('high'), { priority: 100 });
    EventBus.emit('test:priority');
    expect(order).toEqual(['high', 'low']);
  });
});

describe('EventBus parent event matching', () => {
  test('parent event listener fires for child events', () => {
    const handler = jest.fn();
    EventBus.on('sheet:edit', handler);
    EventBus.emit('sheet:edit:GRIEVANCE_LOG', { data: 'test' });
    expect(handler).toHaveBeenCalledWith({ data: 'test' });
  });
});

describe('EventBus.setEnabled', () => {
  test('disabled EventBus does not invoke listeners', () => {
    const handler = jest.fn();
    EventBus.on('test:disabled', handler);
    EventBus.setEnabled(false);
    const result = EventBus.emit('test:disabled');
    expect(handler).not.toHaveBeenCalled();
    expect(result.handled).toBe(0);
    EventBus.setEnabled(true);
  });
});

describe('EventBus.getLog', () => {
  test('logs emitted events', () => {
    EventBus.emit('test:log1');
    EventBus.emit('test:log2');
    const log = EventBus.getLog();
    expect(log.length).toBe(2);
    expect(log[0].event).toBe('test:log1');
    expect(log[1].event).toBe('test:log2');
  });

  test('getLog with count returns limited entries', () => {
    EventBus.emit('a');
    EventBus.emit('b');
    EventBus.emit('c');
    const log = EventBus.getLog(2);
    expect(log.length).toBe(2);
    expect(log[0].event).toBe('b');
  });
});

describe('EventBus.listenerCount', () => {
  test('counts listeners for specific event', () => {
    EventBus.on('test:count', jest.fn());
    EventBus.on('test:count', jest.fn());
    expect(EventBus.listenerCount('test:count')).toBe(2);
  });

  test('counts total listeners across all events', () => {
    EventBus.on('event1', jest.fn());
    EventBus.on('event2', jest.fn());
    EventBus.onAny(jest.fn());
    expect(EventBus.listenerCount()).toBe(3);
  });
});

describe('EventBus.eventNames', () => {
  test('returns registered event names', () => {
    EventBus.on('alpha', jest.fn());
    EventBus.on('beta', jest.fn());
    const names = EventBus.eventNames();
    expect(names).toContain('alpha');
    expect(names).toContain('beta');
  });
});

describe('EventBus.reset', () => {
  test('clears all listeners, wildcards, and log', () => {
    EventBus.on('test', jest.fn());
    EventBus.onAny(jest.fn());
    EventBus.emit('test');
    EventBus.reset();
    expect(EventBus.listenerCount()).toBe(0);
    expect(EventBus.getLog().length).toBe(0);
    expect(EventBus.eventNames().length).toBe(0);
  });
});

// ============================================================================
// Event Bridge Functions
// ============================================================================

describe('emitEditEvent', () => {
  test('returns {handled:0} for null event', () => {
    const result = emitEditEvent(null);
    expect(result.handled).toBe(0);
  });

  test('returns {handled:0} for event without range', () => {
    const result = emitEditEvent({});
    expect(result.handled).toBe(0);
  });

  test('emits sheet:edit:KEY for known sheet names', () => {
    const handler = jest.fn();
    // GRIEVANCE_TRACKER alias removed in v4.25.9 (FIX-CORE-02).
    // emitEditEvent maps 'Grievance Log' -> GRIEVANCE_LOG key.
    EventBus.on('sheet:edit:GRIEVANCE_LOG', handler);

    const mockRange = {
      getSheet: () => ({ getName: () => SHEETS.GRIEVANCE_LOG })
    };
    emitEditEvent({ range: mockRange });
    expect(handler).toHaveBeenCalled();
  });

  test('returns {handled:0} for unknown sheet names', () => {
    const mockRange = {
      getSheet: () => ({ getName: () => 'SomeRandomSheet' })
    };
    const result = emitEditEvent({ range: mockRange });
    expect(result.handled).toBe(0);
  });
});

describe('emitFormEvent', () => {
  test('emits form:submit:type event', () => {
    const handler = jest.fn();
    EventBus.on('form:submit:grievance', handler);
    emitFormEvent('grievance', { id: 'G-001' });
    expect(handler).toHaveBeenCalledWith({ id: 'G-001' });
  });
});

describe('emitSyncComplete', () => {
  test('emits sync:complete event with source', () => {
    const handler = jest.fn();
    EventBus.on('sync:complete', handler);
    emitSyncComplete('grievance');
    expect(handler).toHaveBeenCalledWith(expect.objectContaining({ source: 'grievance' }));
  });
});

describe('emitDataChanged', () => {
  test('emits data:changed event with source and timestamp', () => {
    const handler = jest.fn();
    EventBus.on('data:changed', handler);
    emitDataChanged('member_dir');
    expect(handler).toHaveBeenCalledWith(expect.objectContaining({
      source: 'member_dir',
      timestamp: expect.any(String)
    }));
  });
});

// ============================================================================
// registerEventBusSubscribers
// ============================================================================

describe('registerEventBusSubscribers', () => {
  test('registers subscribers on first call', () => {
    registerEventBusSubscribers();
    expect(EventBus.listenerCount()).toBeGreaterThan(0);
  });

  test('does not re-register on second call', () => {
    registerEventBusSubscribers();
    const countAfterFirst = EventBus.listenerCount();
    registerEventBusSubscribers();
    expect(EventBus.listenerCount()).toBe(countAfterFirst);
  });

  test('registers grievance edit handler', () => {
    registerEventBusSubscribers();
    expect(EventBus.listenerCount('sheet:edit:GRIEVANCE_LOG')).toBeGreaterThan(0);
  });

  test('registers member directory edit handler', () => {
    registerEventBusSubscribers();
    expect(EventBus.listenerCount('sheet:edit:MEMBER_DIR')).toBeGreaterThan(0);
  });

  test('registers config edit handler', () => {
    registerEventBusSubscribers();
    expect(EventBus.listenerCount('sheet:edit:CONFIG')).toBeGreaterThan(0);
  });

  test('registers data changed handler', () => {
    registerEventBusSubscribers();
    expect(EventBus.listenerCount('data:changed')).toBeGreaterThan(0);
  });
});
