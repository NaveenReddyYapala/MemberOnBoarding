import * as React from "react";

export default function FinanceForm({ item }: { item: any }) {
  return (
    <div>
      <h3>Legal Form</h3>
      <p>ID: {item.Id}</p>
      <p>Title: {item.Title}</p>
      <p>Status: {item.Status}</p>
    </div>
  );
}
