import * as React from 'react';
import { useState, useEffect } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import styles from './ApDashboard.module.scss';

export interface IApDashboardProps {
  context: WebPartContext;
}

interface IInvoice {
  ReferenceID: string;
  ClientID: string;
  ClientName: string;
  ProcessingStatus: string;
  SupplierName: string;
  TotalAmountIncGST: number;
  InvoiceDate: string;
  BlobFilePath: string;
  XeroInvoiceURL: { Url: string; Description: string } | string;
}

interface IClient {
  clientId: string;
  ClientName: string;
  apEmailAddress: string;
  status: string;
}

interface IMetrics {
  totalEntities: number;
  invoicesToday: number;
  autoProcessed: number;
  pendingApproval: number;
  completed: number;
  totalValuePending: number;
}

const today = new Date().toISOString().split('T')[0];

export const ApDashboard: React.FC<IApDashboardProps> = ({ context }): React.ReactElement => {
  const [invoices, setInvoices]       = useState<IInvoice[]>([]);
  const [clients, setClients]         = useState<IClient[]>([]);
  const [metrics, setMetrics]         = useState<IMetrics | null>(null);
  const [loading, setLoading]         = useState(true);
  const [error, setError]             = useState<string | null>(null);
  const [lastRefresh, setLastRefresh] = useState<string>('');

  const sp = spfi().using(SPFx(context));

  const fetchData = async (): Promise<void> => {
    try {
      setLoading(true);

      const [invoiceItems, clientItems] = await Promise.all([
        sp.web.lists.getByTitle('InvoiceData').items
          .select('ReferenceID','ClientID','ClientName','ProcessingStatus',
                  'SupplierName','TotalAmountIncGST','InvoiceDate','BlobFilePath','XeroInvoiceURL')
          .top(500)(),
        sp.web.lists.getByTitle('Clients').items
          .select('clientId','ClientName','apEmailAddress','status')
          .top(500)(),
      ]);

      setInvoices(invoiceItems);
      setClients(clientItems);

      const todayInvoices = invoiceItems.filter(i =>
        i.InvoiceDate && i.InvoiceDate.startsWith(today)
      );
      const completed    = invoiceItems.filter(i => i.ProcessingStatus === 'Completed');
      const pending      = invoiceItems.filter(i => i.ProcessingStatus === 'Approval Pending');
      const totalPending = pending.reduce((s, i) => s + (i.TotalAmountIncGST || 0), 0);
      const rate         = invoiceItems.length > 0
        ? Math.round((completed.length / invoiceItems.length) * 100)
        : 0;

      setMetrics({
        totalEntities:     clientItems.length,
        invoicesToday:     todayInvoices.length,
        autoProcessed:     rate,
        pendingApproval:   pending.length,
        completed:         completed.length,
        totalValuePending: totalPending,
      });

      setLastRefresh(new Date().toLocaleTimeString('en-AU', {
        hour: '2-digit', minute: '2-digit', timeZoneName: 'short'
      }));
      setError(null);
    } catch (e) {
      setError('Failed to load data. ' + (e as Error).message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    fetchData();
    const interval = setInterval(() => { fetchData().catch(console.error); }, 60000);
    return () => clearInterval(interval);
  }, []);

  const getClientRate = (clientId: string): number => {
    const clientInvoices = invoices.filter(i => i.ClientID === clientId);
    if (clientInvoices.length === 0) return 0;
    const done = clientInvoices.filter(i => i.ProcessingStatus === 'Completed').length;
    return Math.round((done / clientInvoices.length) * 100);
  };

  const getClientStatusLabel = (rate: number): string => {
    if (rate >= 85) return 'Completed';
    if (rate >= 70) return 'In Progress';
    return 'Pending';
  };

  const getClientStatusClass = (rate: number): string => {
    if (rate >= 85) return styles.sOk;
    if (rate >= 70) return styles.sWarn;
    return styles.sBlock;
  };

  const getBarColor = (rate: number): string => {
    if (rate >= 85) return '#1B3A6B';
    if (rate >= 70) return '#BA7517';
    return '#A32D2D';
  };

  const attentionItems = invoices.filter(i =>
    i.ProcessingStatus === 'Approval Pending'
  ).slice(0, 10);

  if (loading) return (
    <div className={styles.loading}>
      <div className={styles.spinner} />
      <span>Loading AP Dashboard...</span>
    </div>
  );

  if (error) return (
    <div className={styles.error}>{error}</div>
  );

  return (
    <div className={styles.root}>
      {/* NAV */}
      <div className={styles.nav}>
        <span className={styles.navLogo}>Vanderbilt Equity</span>
        <span className={styles.navSep}>|</span>
        <span className={styles.navSite}>Managed AP Platform</span>
        <div className={styles.navRight}>
          <span className={styles.dot} />
          Live · Auto-refreshes every 60s · Last updated {lastRefresh}
        </div>
      </div>

      {/* HEADER */}
      <div className={styles.hdr}>
        <div>
          <div className={styles.hdrTitle}>AP Operations Dashboard</div>
          <div className={styles.hdrSub}>
            {metrics?.totalEntities} active entities · All Xero accounts connected · Real-time view
          </div>
        </div>
        <div className={styles.hdrDate}>
          {new Date().toLocaleDateString('en-AU', { weekday:'short', day:'numeric', month:'short', year:'numeric' })}
          &nbsp;·&nbsp;{lastRefresh}
        </div>
      </div>

      <div className={styles.body}>

        {/* METRICS */}
        <div className={styles.metrics}>
          <div className={styles.mc}>
            <div className={styles.mcLbl}>Active entities</div>
            <div className={styles.mcVal}>{metrics?.totalEntities}</div>
            <div className={styles.mcSub}>All Xero connected</div>
          </div>
          <div className={styles.mc}>
            <div className={styles.mcLbl}>Invoices today</div>
            <div className={styles.mcVal}>{metrics?.invoicesToday}</div>
            <div className={styles.mcSub}>{metrics?.completed} completed</div>
          </div>
          <div className={styles.mc}>
            <div className={styles.mcLbl}>Auto-process rate</div>
            <div className={styles.mcVal}>{metrics?.autoProcessed}%</div>
            <div className={styles.mcSub}>Target: 95%</div>
          </div>
          <div className={styles.mc}>
            <div className={styles.mcLbl}>Pending approval</div>
            <div className={styles.mcVal}>{metrics?.pendingApproval}</div>
            <div className={styles.mcSub}>
              ${metrics?.totalValuePending.toLocaleString('en-AU', { maximumFractionDigits: 0 })} total value
            </div>
          </div>
          <div className={styles.mc}>
            <div className={styles.mcLbl}>Completed</div>
            <div className={styles.mcVal}>{metrics?.completed}</div>
            <div className={styles.mcSub}>Posted to Xero</div>
          </div>
        </div>

        {/* ENTITY GRID */}
        <div className={styles.sec}>All entities — live invoice status</div>
        <div className={styles.entGrid}>
          {clients.map(client => {
            const rate  = getClientRate(client.clientId);
            const color = getBarColor(rate);
            const label = getClientStatusLabel(rate);
            const cls   = getClientStatusClass(rate);
            return (
              <div key={client.clientId} className={styles.ent}>
                <div className={styles.entName}>{client.ClientName}</div>
                <div className={styles.entEmail}>{client.apEmailAddress}</div>
                <div className={styles.barBg}>
                  <div className={styles.barFill} style={{ width: `${rate}%`, background: color }} />
                </div>
                <div className={`${styles.entStatus} ${cls}`}>{label} — {rate}%</div>
              </div>
            );
          })}
        </div>

        {/* ATTENTION TABLE */}
        <div className={styles.sec}>Items requiring attention</div>
        <div className={styles.tblWrap}>
          <table className={styles.tbl}>
            <thead>
              <tr>
                <th>Entity</th>
                <th>Status</th>
                <th>Supplier</th>
                <th>Amount</th>
                <th>Invoice date</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {attentionItems.length === 0
                ? <tr><td colSpan={6} style={{ textAlign: 'center', color: '#888', padding: '20px' }}>No items requiring attention</td></tr>
                : attentionItems.map(inv => (
                  <tr key={inv.ReferenceID}>
                    <td>{inv.ClientName}</td>
                    <td>
                      <span className={`${styles.bdg} ${styles.bWarn}`}>
                        {inv.ProcessingStatus}
                      </span>
                    </td>
                    <td>{inv.SupplierName}</td>
                    <td>${(inv.TotalAmountIncGST || 0).toLocaleString('en-AU', { minimumFractionDigits: 2 })}</td>
                    <td>{inv.InvoiceDate ? new Date(inv.InvoiceDate).toLocaleDateString('en-AU') : '—'}</td>
                    <td>
                      <div style={{ display: 'flex', gap: '16px', alignItems: 'center' }}>
                        {inv.BlobFilePath
                          ? <a href={inv.BlobFilePath} target="_blank" rel="noreferrer" className={styles.act}>View file</a>
                          : <span className={styles.actDisabled}>No file</span>
                        }
                        {inv.XeroInvoiceURL
                          ? <a href={typeof inv.XeroInvoiceURL === 'object' ? inv.XeroInvoiceURL.Url : inv.XeroInvoiceURL} target="_blank" rel="noreferrer" className={styles.actXero}>View in Xero</a>
                          : <span className={styles.actDisabled}>—</span>
                        }
                      </div>
                    </td>
                  </tr>
                ))
              }
            </tbody>
          </table>
        </div>

        {/* SUMMARY PANELS */}
        <div className={styles.panels}>
          <div className={styles.panel}>
            <div className={styles.panelTitle}>Invoice summary — all time</div>
            {[
              ['Total invoices', invoices.length],
              ['Completed / posted to Xero', invoices.filter(i => i.ProcessingStatus === 'Completed').length],
              ['Pending approval', invoices.filter(i => i.ProcessingStatus === 'Approval Pending').length],
              ['In progress', invoices.filter(i => i.ProcessingStatus === 'In Progress').length],
              ['New', invoices.filter(i => i.ProcessingStatus === 'New').length],
            ].map(([label, val]) => (
              <div key={label as string} className={styles.sr}>
                <span className={styles.srName}>{label}</span>
                <span className={styles.srVal}>{val}</span>
              </div>
            ))}
          </div>
          <div className={styles.panel}>
            <div className={styles.panelTitle}>Processing rate by status</div>
            {[
              { label: 'Completed', count: invoices.filter(i => i.ProcessingStatus === 'Completed').length, color: '#1B3A6B' },
              { label: 'In Progress', count: invoices.filter(i => i.ProcessingStatus === 'In Progress').length, color: '#1B3A6B' },
              { label: 'Approval Pending', count: invoices.filter(i => i.ProcessingStatus === 'Approval Pending').length, color: '#BA7517' },
              { label: 'New', count: invoices.filter(i => i.ProcessingStatus === 'New').length, color: '#1B3A6B' },
            ].map(({ label, count, color }) => {
              const pct = invoices.length > 0 ? Math.round((count / invoices.length) * 100) : 0;
              return (
                <div key={label} className={styles.br}>
                  <div className={styles.brLbl}>
                    <span>{label}</span>
                    <span className={styles.brVal}>{pct}% ({count})</span>
                  </div>
                  <div className={styles.barTrack}>
                    <div className={styles.barInner} style={{ width: `${pct}%`, background: color }} />
                  </div>
                </div>
              );
            })}
          </div>
        </div>

      </div>

      {/* FOOTER */}
      <div className={styles.footer}>
        <span className={styles.fl}>Vanderbilt Equity · Managed AP Platform · Confidential</span>
        <span className={styles.fl}>Powered by Business Automation Service · businessautomationservice.com</span>
      </div>
    </div>
  );
};