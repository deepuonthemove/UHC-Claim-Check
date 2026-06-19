'use client';

import { useState, useRef, useCallback } from 'react';
import { fetchEventSource } from '@microsoft/fetch-event-source';
import ExcelJS from 'exceljs';
import styles from './ProgressLog.module.css';

interface LogEntry {
  id: number;
  type: 'log' | 'progress' | 'error' | 'success' | 'info';
  message: string;
  timestamp: string;
}

interface ProgressLogProps {
  entries: LogEntry[];
}

export function ProgressLog({ entries }: ProgressLogProps) {
  const bottomRef = useRef<HTMLDivElement>(null);
  return (
    <div className={styles.logContainer}>
      <div className={styles.logHeader}>
        <span className={styles.logTitle}>Live Log</span>
        <span className={styles.logCount}>{entries.length} events</span>
      </div>
      <div className={styles.logBody}>
        {entries.length === 0 && (
          <div className={styles.logEmpty}>Waiting for automation to start…</div>
        )}
        {entries.map(entry => (
          <div key={entry.id} className={`${styles.logEntry} ${styles[`logEntry_${entry.type}`]}`}>
            <span className={styles.logTime}>{entry.timestamp}</span>
            <span className={styles.logMsg}>{entry.message}</span>
          </div>
        ))}
        <div ref={bottomRef} />
      </div>
    </div>
  );
}

export type { LogEntry };
