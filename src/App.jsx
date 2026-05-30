import React from "react";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer
} from "recharts";

const kpis = {
  totalWO: 500,
  totalInvoice: 973064,
  avgInvoice: 1946,
  providers: 38
};

const spendByTrade = [
  { name: "HANDYMAN", value: 428605 },
  { name: "CLEANING", value: 403839 },
  { name: "PM GENERAL MAINTENANCE", value: 115725 },
  { name: "HVAC", value: 19871 },
  { name: "ELECTRICAL", value: 2453 }
];

function Card({ label, value }) {
  return (
    <div className="card">
      <p className="label">{label}</p>
      <h2>{value}</h2>
    </div>
  );
}

export default function App() {
  return (
    <main className="page">
      <h1>FM Executive Dashboard</h1>
      <p className="subtitle">
        Executive view of work orders, invoice amount, providers and trades.
      </p>

      <section className="grid">
        <Card label="Work Orders" value={kpis.totalWO.toLocaleString()} />
        <Card label="Invoice Amount" value={`€${kpis.totalInvoice.toLocaleString()}`} />
        <Card label="Avg. Invoice / WO" value={`€${kpis.avgInvoice.toLocaleString()}`} />
        <Card label="Providers" value={kpis.providers} />
      </section>

      <section className="panel">
        <p className="label">Cost concentration</p>
        <h2>Spend by Trade</h2>

        <div style={{ width: "100%", height: 420 }}>
          <ResponsiveContainer>
            <BarChart data={spendByTrade} layout="vertical">
              <XAxis type="number" stroke="#aaa" />
              <YAxis dataKey="name" type="category" width={180} stroke="#f7f3ea" />
              <Tooltip />
              <Bar dataKey="value" fill="#d6b56d" radius={[0, 8, 8, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </section>
    </main>
  );
}
