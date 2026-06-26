import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid,
  PieChart,
  Pie,
  Cell
} from "recharts";
import {
  Briefcase,
  Building2,
  Euro,
  Users,
  AlertTriangle,
  Activity
} from "lucide-react";

const FILE_URL = "/fm-data.xlsx";

function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  return Number(String(value).replace(",", ".").replace(/[^\d.-]/g, "")) || 0;
}

function formatEuro(value) {
  return new Intl.NumberFormat("es-ES", {
    style: "currency",
    currency: "EUR",
    maximumFractionDigits: 0
  }).format(value);
}

function topBy(rows, groupCol, valueCol, limit = 10) {
  const map = {};

  rows.forEach((row) => {
    const key = row[groupCol] || "No informado";
    const value = toNumber(row[valueCol]);
    map[key] = (map[key] || 0) + value;
  });

  return Object.entries(map)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value)
    .slice(0, limit);
}

function countBy(rows, groupCol) {
  const map = {};

  rows.forEach((row) => {
    const key = row[groupCol] || "No informado";
    map[key] = (map[key] || 0) + 1;
  });

  return Object.entries(map)
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);
}

function KpiCard({ title, value, subtitle, icon: Icon }) {
  return (
    <div className="kpi-card">
      <div>
        <p className="kpi-title">{title}</p>
        <h2>{value}</h2>
        <p className="kpi-subtitle">{subtitle}</p>
      </div>
      <div className="kpi-icon">
        <Icon size={22} />
      </div>
    </div>
  );
}

export default function App() {
  const [rows, setRows] = useState([]);
  const [status, setStatus] = useState("Cargando Excel...");

  useEffect(() => {
    async function loadExcel() {
      try {
        const response = await fetch(FILE_URL);
        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const firstSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheet];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        setRows(json);
        setStatus(`Excel cargado correctamente: ${json.length} registros`);
      } catch (error) {
        console.error(error);
        setStatus("Error cargando el Excel");
      }
    }

    loadExcel();
  }, []);

  const metrics = useMemo(() => {
    const totalWO = rows.length;
    const totalSpend = rows.reduce((sum, row) => sum + toNumber(row["Inv. Amount"]), 0);
    const totalNTE = rows.reduce((sum, row) => sum + toNumber(row["NTE"]), 0);

    const providers = new Set(rows.map((r) => r["Provider Name"]).filter(Boolean)).size;
    const locations = new Set(rows.map((r) => r["Location Number"]).filter(Boolean)).size;

    const avgInvoice = totalWO ? totalSpend / totalWO : 0;

    const completed = rows.filter((r) => r["WO Status"] === "COMPLETED").length;
    const preventive = rows.filter((r) => r["Category"] === "PREVENTIVE").length;
    const corrective = rows.filter((r) => r["Category"] === "CORRECTIVE").length;

    const overNTE = rows.filter((r) => {
      const inv = toNumber(r["Inv. Amount"]);
      const nte = toNumber(r["NTE"]);
      return nte > 0 && inv > nte;
    }).length;

    return {
      totalWO,
      totalSpend,
      totalNTE,
      providers,
      locations,
      avgInvoice,
      completed,
      preventive,
      corrective,
      overNTE
    };
  }, [rows]);

  const spendByTrade = useMemo(
    () => topBy(rows, "Trade", "Inv. Amount", 10),
    [rows]
  );

  const spendByProvider = useMemo(
    () => topBy(rows, "Provider Name", "Inv. Amount", 10),
    [rows]
  );

  const priorityMix = useMemo(
    () => countBy(rows, "Priority").slice(0, 6),
    [rows]
  );

  const statusMix = useMemo(
    () => countBy(rows, "WO Status").slice(0, 8),
    [rows]
  );

  const topWO = useMemo(() => {
    return [...rows]
      .sort((a, b) => toNumber(b["Inv. Amount"]) - toNumber(a["Inv. Amount"]))
      .slice(0, 10);
  }, [rows]);

  const pieColors = ["#EB8C00", "#DB536A", "#464646", "#FFB600", "#7D7D7D", "#E0301E"];

  return (
    <main className="page">
      <header className="hero">
        <div>
          <p className="eyebrow">PwC style · FM Analytics</p>
          <h1>Executive Facility Management Dashboard</h1>
          <p className="hero-text">
            Financial and operational analysis of work orders, invoices, providers,
            trades, priorities and risk signals.
          </p>
        </div>

        <div className="status-box">
          <span>{status}</span>
        </div>
      </header>

      <section className="kpi-grid">
        <KpiCard
          title="Total Spend"
          value={formatEuro(metrics.totalSpend)}
          subtitle="Total invoice amount"
          icon={Euro}
        />
        <KpiCard
          title="Work Orders"
          value={metrics.totalWO.toLocaleString("es-ES")}
          subtitle="Total records loaded"
          icon={Briefcase}
        />
        <KpiCard
          title="Avg. Invoice / WO"
          value={formatEuro(metrics.avgInvoice)}
          subtitle="Average cost per work order"
          icon={Activity}
        />
        <KpiCard
          title="Providers"
          value={metrics.providers}
          subtitle="Unique providers"
          icon={Users}
        />
        <KpiCard
          title="Locations"
          value={metrics.locations}
          subtitle="Unique locations"
          icon={Building2}
        />
        <KpiCard
          title="Over NTE"
          value={metrics.overNTE}
          subtitle="Invoices above NTE"
          icon={AlertTriangle}
        />
      </section>

      <section className="insight-panel">
        <p className="eyebrow">Executive reading</p>
        <h3>Key signals from the current file</h3>
        <ul>
          <li>
            The file contains <strong>{metrics.totalWO.toLocaleString("es-ES")}</strong> work orders
            across <strong>{metrics.locations}</strong> locations.
          </li>
          <li>
            Total invoice amount is <strong>{formatEuro(metrics.totalSpend)}</strong>, with an
            average cost of <strong>{formatEuro(metrics.avgInvoice)}</strong> per work order.
          </li>
          <li>
            The dashboard has detected <strong>{metrics.overNTE}</strong> cases where invoice
            amount is above NTE.
          </li>
        </ul>
      </section>

      <section className="two-columns">
        <div className="panel">
          <p className="eyebrow">Cost concentration</p>
          <h3>Spend by Trade</h3>
          <div className="chart">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={spendByTrade} layout="vertical">
                <CartesianGrid stroke="rgba(0,0,0,0.08)" horizontal={false} />
                <XAxis type="number" tickFormatter={(v) => `€${Math.round(v / 1000)}k`} />
                <YAxis dataKey="name" type="category" width={160} />
                <Tooltip formatter={(value) => formatEuro(value)} />
                <Bar dataKey="value" fill="#EB8C00" radius={[0, 8, 8, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="panel">
          <p className="eyebrow">Supplier concentration</p>
          <h3>Spend by Provider</h3>
          <div className="chart">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={spendByProvider} layout="vertical">
                <CartesianGrid stroke="rgba(0,0,0,0.08)" horizontal={false} />
                <XAxis type="number" tickFormatter={(v) => `€${Math.round(v / 1000)}k`} />
                <YAxis dataKey="name" type="category" width={170} />
                <Tooltip formatter={(value) => formatEuro(value)} />
                <Bar dataKey="value" fill="#E0301E" radius={[0, 8, 8, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </section>

      <section className="two-columns">
        <div className="panel">
          <p className="eyebrow">Operational urgency</p>
          <h3>Priority Mix</h3>
          <div className="chart">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={priorityMix}
                  dataKey="value"
                  nameKey="name"
                  innerRadius={70}
                  outerRadius={120}
                  paddingAngle={4}
                >
                  {priorityMix.map((entry, index) => (
                    <Cell key={entry.name} fill={pieColors[index % pieColors.length]} />
                  ))}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="panel">
          <p className="eyebrow">Execution status</p>
          <h3>WO Status Mix</h3>
          <div className="chart">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={statusMix}>
                <CartesianGrid stroke="rgba(0,0,0,0.08)" vertical={false} />
                <XAxis dataKey="name" />
                <YAxis />
                <Tooltip />
                <Bar dataKey="value" fill="#464646" radius={[8, 8, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </section>

      <section className="panel">
        <p className="eyebrow">Risk review</p>
        <h3>Top 10 Most Expensive Work Orders</h3>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>WO</th>
                <th>Date</th>
                <th>Location</th>
                <th>Trade</th>
                <th>Category</th>
                <th>Priority</th>
                <th>Provider</th>
                <th>NTE</th>
                <th>Invoice</th>
              </tr>
            </thead>
            <tbody>
              {topWO.map((row, index) => (
                <tr key={index}>
                  <td>{row["WO Tracking Number"]}</td>
                  <td>{String(row["Call Date"]).slice(0, 10)}</td>
                  <td>{row["Location Number"]}</td>
                  <td>{row["Trade"]}</td>
                  <td>{row["Category"]}</td>
                  <td>{row["Priority"]}</td>
                  <td>{row["Provider Name"]}</td>
                  <td>{formatEuro(toNumber(row["NTE"]))}</td>
                  <td className="strong">{formatEuro(toNumber(row["Inv. Amount"]))}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </main>
  );
}
