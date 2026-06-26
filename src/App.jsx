import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid
} from "recharts";

const FILE_URL = "/fm-data.xlsx";

const COUNTRY_COST_FACTOR = {
  France: 1.0,
  Spain: 0.9,
  Italy: 0.95,
  Germany: 1.15,
  Austria: 1.12,
  Switzerland: 1.45,
  Netherlands: 1.18,
  Belgium: 1.12,
  UK: 1.2,
  Portugal: 0.82,
  Poland: 0.72
};

function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  return Number(String(value).replace(",", ".").replace(/[^\d.-]/g, "")) || 0;
}

function formatEuro(value) {
  return new Intl.NumberFormat("es-ES", {
    style: "currency",
    currency: "EUR",
    maximumFractionDigits: 0
  }).format(value || 0);
}

function percentileRank(values, value) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const below = sorted.filter((v) => v <= value).length;
  return below / sorted.length;
}

function groupBy(rows, keyFn) {
  const map = {};
  rows.forEach((row) => {
    const key = keyFn(row) || "No informado";
    if (!map[key]) map[key] = [];
    map[key].push(row);
  });
  return map;
}

function buildStoreMaster(workbook) {
  const possibleSheet = workbook.SheetNames.find((name) =>
    ["STORE_MASTER", "LOCATION_MASTER", "MASTER", "STORES"].includes(
      name.toUpperCase()
    )
  );

  if (!possibleSheet) return {};

  const ws = workbook.Sheets[possibleSheet];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

  const map = {};

  rows.forEach((row) => {
    const location = String(row["Location Number"] || "").trim();
    const country = String(row["Country"] || "").trim();

    if (location && country) {
      map[location] = country;
    }
  });

  return map;
}

function enrichRowsWithCountry(rows, storeMaster) {
  return rows.map((row) => {
    const location = String(row["Location Number"] || "").trim();

    return {
      ...row,
      Country:
        row["Country"] ||
        storeMaster[location] ||
        "Unknown"
    };
  });
}

function buildCountryBenchmark(rows) {
  const byCountry = groupBy(rows, (row) => row.Country);
  const result = Object.entries(byCountry).map(([country, countryRows]) => {
    const wo = countryRows.length;
    const spend = countryRows.reduce(
      (sum, row) => sum + toNumber(row["Inv. Amount"]),
      0
    );
    const avgCost = wo ? spend / wo : 0;

    return {
      country,
      wo,
      spend,
      avgCost
    };
  });

  const france = result.find((r) => r.country === "France");
  const franceAvg = france?.avgCost || result[0]?.avgCost || 1;

  return result
    .map((row) => {
      const factor = COUNTRY_COST_FACTOR[row.country] || 1;
      const expectedAvg = franceAvg * factor;
      const franceIndex = franceAvg ? (row.avgCost / franceAvg) * 100 : 0;
      const cpi = expectedAvg ? row.avgCost / expectedAvg : 0;

      return {
        ...row,
        factor,
        expectedAvg,
        franceIndex,
        cpi
      };
    })
    .sort((a, b) => b.franceIndex - a.franceIndex);
}

function buildStoreRanking(rows) {
  const byStore = groupBy(rows, (row) => String(row["Location Number"] || ""));

  const stores = Object.entries(byStore).map(([location, storeRows]) => {
    const wo = storeRows.length;
    const spend = storeRows.reduce(
      (sum, row) => sum + toNumber(row["Inv. Amount"]),
      0
    );
    const avgInvoice = wo ? spend / wo : 0;

    const emergency = storeRows.filter((row) =>
      String(row["Priority"] || "").toUpperCase().includes("RED")
    ).length;

    const corrective = storeRows.filter((row) =>
      String(row["Category"] || "").toUpperCase().includes("CORRECTIVE")
    ).length;

    const overNTE = storeRows.filter((row) => {
      const nte = toNumber(row["NTE"]);
      const inv = toNumber(row["Inv. Amount"]);
      return nte > 0 && inv > nte;
    }).length;

    const providers = new Set(
      storeRows.map((row) => row["Provider Name"]).filter(Boolean)
    ).size;

    const country = storeRows[0]?.Country || "Unknown";

    return {
      location,
      country,
      wo,
      spend,
      avgInvoice,
      emergency,
      corrective,
      overNTE,
      providers,
      emergencyRate: wo ? emergency / wo : 0,
      correctiveRate: wo ? corrective / wo : 0,
      overNTERate: wo ? overNTE / wo : 0
    };
  });

  const spendValues = stores.map((s) => s.spend);
  const woValues = stores.map((s) => s.wo);
  const avgValues = stores.map((s) => s.avgInvoice);
  const emergencyValues = stores.map((s) => s.emergencyRate);
  const overNTEValues = stores.map((s) => s.overNTERate);
  const providerValues = stores.map((s) => s.providers);

  return stores
    .map((store) => {
      const financialRisk =
        percentileRank(spendValues, store.spend) * 0.28 +
        percentileRank(avgValues, store.avgInvoice) * 0.22;

      const operationalRisk =
        percentileRank(woValues, store.wo) * 0.2 +
        percentileRank(emergencyValues, store.emergencyRate) * 0.12 +
        percentileRank(overNTEValues, store.overNTERate) * 0.12 +
        percentileRank(providerValues, store.providers) * 0.06;

      const riskScore = Math.round((financialRisk + operationalRisk) * 100);
      const healthScore = Math.max(0, 100 - riskScore);

      return {
        ...store,
        riskScore,
        healthScore
      };
    })
    .sort((a, b) => b.riskScore - a.riskScore);
}

function KpiCard({ title, value, subtitle }) {
  return (
    <div className="kpi-card">
      <p className="kpi-title">{title}</p>
      <h2>{value}</h2>
      <p className="kpi-subtitle">{subtitle}</p>
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

        const mainSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[mainSheet];

        const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        const storeMaster = buildStoreMaster(workbook);
        const enrichedRows = enrichRowsWithCountry(rawRows, storeMaster);

        setRows(enrichedRows);
        setStatus(`Excel cargado correctamente: ${enrichedRows.length} registros`);
      } catch (error) {
        console.error(error);
        setStatus("Error cargando el Excel");
      }
    }

    loadExcel();
  }, []);

  const metrics = useMemo(() => {
    const totalWO = rows.length;
    const totalSpend = rows.reduce(
      (sum, row) => sum + toNumber(row["Inv. Amount"]),
      0
    );

    const providers = new Set(
      rows.map((r) => r["Provider Name"]).filter(Boolean)
    ).size;

    const locations = new Set(
      rows.map((r) => r["Location Number"]).filter(Boolean)
    ).size;

    const countries = new Set(
      rows.map((r) => r["Country"]).filter((c) => c && c !== "Unknown")
    ).size;

    const avgInvoice = totalWO ? totalSpend / totalWO : 0;

    return {
      totalWO,
      totalSpend,
      providers,
      locations,
      countries,
      avgInvoice
    };
  }, [rows]);

  const countryBenchmark = useMemo(
    () => buildCountryBenchmark(rows).filter((r) => r.country !== "Unknown"),
    [rows]
  );

  const storeRanking = useMemo(() => buildStoreRanking(rows), [rows]);

  const topRiskStores = storeRanking.slice(0, 15);

  return (
    <main className="page">
      <header className="hero">
        <div>
          <p className="eyebrow">PwC style · FM Analytics</p>
          <h1>Executive Facility Management Dashboard</h1>
          <p className="hero-text">
            Cost benchmarking, store ranking and financial risk analysis based on work order data.
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
        />
        <KpiCard
          title="Work Orders"
          value={metrics.totalWO.toLocaleString("es-ES")}
          subtitle="Total records loaded"
        />
        <KpiCard
          title="Avg. Invoice / WO"
          value={formatEuro(metrics.avgInvoice)}
          subtitle="Average cost per work order"
        />
        <KpiCard
          title="Providers"
          value={metrics.providers}
          subtitle="Unique providers"
        />
        <KpiCard
          title="Locations"
          value={metrics.locations}
          subtitle="Unique stores"
        />
        <KpiCard
          title="Countries"
          value={metrics.countries}
          subtitle="Countries detected"
        />
      </section>

      {countryBenchmark.length === 0 && (
        <section className="warning-panel">
          <p className="eyebrow">Country benchmark unavailable</p>
          <h3>Falta información de país</h3>
          <p>
            El fichero actual no contiene una columna <strong>Country</strong>.
            Para activar el benchmark mundial, añade una segunda pestaña llamada{" "}
            <strong>STORE_MASTER</strong> con estas columnas:
          </p>
          <pre>Location Number | Country</pre>
        </section>
      )}

      {countryBenchmark.length > 0 && (
        <section className="panel">
          <p className="eyebrow">Country benchmark</p>
          <h3>Average cost by country indexed to France = 100</h3>
          <div className="chart">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={countryBenchmark.slice(0, 15)} layout="vertical">
                <CartesianGrid stroke="rgba(0,0,0,0.08)" horizontal={false} />
                <XAxis type="number" />
                <YAxis dataKey="country" type="category" width={150} />
                <Tooltip
                  formatter={(value, name) => {
                    if (name === "franceIndex") return [`${Math.round(value)}`, "Index"];
                    return value;
                  }}
                />
                <Bar
                  dataKey="franceIndex"
                  fill="#EB8C00"
                  radius={[0, 8, 8, 0]}
                />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </section>
      )}

      {countryBenchmark.length > 0 && (
        <section className="panel">
          <p className="eyebrow">Cost Performance Index</p>
          <h3>Real cost vs expected country cost</h3>

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Country</th>
                  <th>WO</th>
                  <th>Total Spend</th>
                  <th>Avg. Cost</th>
                  <th>France Index</th>
                  <th>Country Factor</th>
                  <th>Expected Avg.</th>
                  <th>CPI</th>
                </tr>
              </thead>
              <tbody>
                {countryBenchmark.map((row) => (
                  <tr key={row.country}>
                    <td>{row.country}</td>
                    <td>{row.wo}</td>
                    <td>{formatEuro(row.spend)}</td>
                    <td>{formatEuro(row.avgCost)}</td>
                    <td>{Math.round(row.franceIndex)}</td>
                    <td>{row.factor}</td>
                    <td>{formatEuro(row.expectedAvg)}</td>
                    <td className={row.cpi > 1.15 ? "bad" : row.cpi < 0.9 ? "good" : ""}>
                      {row.cpi.toFixed(2)}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      <section className="panel">
        <p className="eyebrow">Store ranking</p>
        <h3>Top stores by risk score</h3>
        <div className="chart">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={topRiskStores} layout="vertical">
              <CartesianGrid stroke="rgba(0,0,0,0.08)" horizontal={false} />
              <XAxis type="number" />
              <YAxis dataKey="location" type="category" width={120} />
              <Tooltip />
              <Bar dataKey="riskScore" fill="#E0301E" radius={[0, 8, 8, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </section>

      <section className="panel">
        <p className="eyebrow">Store intelligence</p>
        <h3>Numerical ranking by store</h3>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Rank</th>
                <th>Location</th>
                <th>Country</th>
                <th>Risk Score</th>
                <th>Health Score</th>
                <th>WO</th>
                <th>Total Spend</th>
                <th>Avg. Invoice</th>
                <th>Emergency</th>
                <th>Over NTE</th>
                <th>Providers</th>
              </tr>
            </thead>
            <tbody>
              {storeRanking.slice(0, 50).map((store, index) => (
                <tr key={store.location}>
                  <td>{index + 1}</td>
                  <td>{store.location}</td>
                  <td>{store.country}</td>
                  <td className="bad">{store.riskScore}</td>
                  <td className={store.healthScore > 70 ? "good" : store.healthScore < 40 ? "bad" : ""}>
                    {store.healthScore}
                  </td>
                  <td>{store.wo}</td>
                  <td>{formatEuro(store.spend)}</td>
                  <td>{formatEuro(store.avgInvoice)}</td>
                  <td>{store.emergency}</td>
                  <td>{store.overNTE}</td>
                  <td>{store.providers}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </main>
  );
}
