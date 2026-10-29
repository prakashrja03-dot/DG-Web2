import React, { useEffect, useState } from "react";
import { useMsal } from "@azure/msal-react";
import { graphConfig } from "./graphConfig";
import {
  getAccessToken,
  getSiteId,
  getListId,
  getListItems,
  updateListItem
} from "./graphService";

const COLUMNS = [
  "EmployeeCode", "EmployeeName", "Email", "Division",
  "CostCenter", "Department", "Designation", "Role", "ReportingManager"
];

function EmployeeList() {
  const { instance, accounts } = useMsal();
  const [siteId, setSiteId] = useState(null);
  const [listId, setListId] = useState(null);
  const [rows, setRows] = useState([]);
  const [editingId, setEditingId] = useState(null);
  const [editValues, setEditValues] = useState({});
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    async function load() {
      try {
        const token = await getAccessToken(instance, accounts);
        const sId = await getSiteId(token, graphConfig.siteHostname, graphConfig.sitePath);
        const lId = await getListId(token, sId, graphConfig.listName);
        const items = await getListItems(token, sId, lId);
        setSiteId(sId);
        setListId(lId);
        setRows(items);
      } catch (err) {
        setError(err.message);
      } finally {
        setLoading(false);
      }
    }
    load();
  }, [instance, accounts]);

  function startEdit(row) {
    setEditingId(row.id);
    setEditValues({ ...row });
  }

  function cancelEdit() {
    setEditingId(null);
    setEditValues({});
  }

  async function saveEdit(id) {
    try {
      const token = await getAccessToken(instance, accounts);
      const fieldsToUpdate = {};
      COLUMNS.forEach((col) => (fieldsToUpdate[col] = editValues[col] || ""));
      await updateListItem(token, siteId, listId, id, fieldsToUpdate);
      setRows((prev) => prev.map((r) => (r.id === id ? { ...r, ...fieldsToUpdate } : r)));
      setEditingId(null);
    } catch (err) {
      alert("Update failed: " + err.message);
    }
  }

  if (loading) return <p style={styles.msg}>Loading employees...</p>;
  if (error) return <p style={{ ...styles.msg, color: "#d92626" }}>Error: {error}</p>;

  return (
    <div style={styles.wrapper}>
      <h2 style={styles.heading}>Employee Masters</h2>
      <div style={styles.tableContainer}>
        <table style={styles.table}>
          <thead>
            <tr>
              {COLUMNS.map((col) => (
                <th key={col} style={styles.th}>{col}</th>
              ))}
              <th style={styles.th}>Action</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={row.id}>
                {COLUMNS.map((col) => (
                  <td key={col} style={styles.td}>
                    {editingId === row.id ? (
                      <input
                        style={styles.input}
                        value={editValues[col] || ""}
                        onChange={(e) =>
                          setEditValues({ ...editValues, [col]: e.target.value })
                        }
                      />
                    ) : (
                      row[col] || ""
                    )}
                  </td>
                ))}
                <td style={styles.td}>
                  {editingId === row.id ? (
                    <>
                      <button style={styles.btnSave} onClick={() => saveEdit(row.id)}>Save</button>
                      <button style={styles.btnCancel} onClick={cancelEdit}>Cancel</button>
                    </>
                  ) : (
                    <button style={styles.btnEdit} onClick={() => startEdit(row)}>Edit</button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

const styles = {
  wrapper: { padding: "20px" },
  heading: { marginBottom: "14px", color: "#1f2937", fontSize: "20px" },
  tableContainer: { overflowX: "auto", border: "1px solid #e5e7eb", borderRadius: "8px" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: "13px" },
  th: {
    background: "#f3f4f6", textAlign: "left", padding: "10px",
    borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap"
  },
  td: { padding: "8px 10px", borderBottom: "1px solid #f1f1f1", whiteSpace: "nowrap" },
  input: { width: "100%", padding: "4px 6px", fontSize: "13px", border: "1px solid #d1d5db", borderRadius: "4px" },
  btnEdit: { background: "#0f6cbd", color: "#fff", border: "none", padding: "5px 10px", borderRadius: "4px", cursor: "pointer", fontSize: "12px" },
  btnSave: { background: "#16a34a", color: "#fff", border: "none", padding: "5px 10px", borderRadius: "4px", cursor: "pointer", fontSize: "12px", marginRight: "6px" },
  btnCancel: { background: "#fff", color: "#6b7280", border: "1px solid #d1d5db", padding: "5px 10px", borderRadius: "4px", cursor: "pointer", fontSize: "12px" },
  msg: { padding: "20px", textAlign: "center", color: "#6b7280" }
};

export default EmployeeList;
