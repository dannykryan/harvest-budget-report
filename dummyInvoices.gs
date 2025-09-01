function getDummyinvoices() {
    const invoicesArray = [
      {
        id: 48124922,
        number: "727",
        amount: 960.0,
        subject: "Website Redesign annual hosting fee",
        state: "open",
        issue_date: "2025-09-01",
        due_date: "2025-10-31",
        currency: "GBP",
        client: { id: 20001, name: "Acme Corp" },
        project_id: 100001,
      },
      {
        id: 48481677,
        number: "766",
        amount: 4750.0,
        subject: "Overbudget Project 2 Pen Testing Actions",
        state: "open",
        issue_date: "2025-08-29",
        due_date: "2025-09-28",
        currency: "GBP",
        client: { id: 20011, name: "OverCorp" },
        project_id: 100011,
      },
      {
        id: 48478824,
        number: "765",
        amount: 6324.0,
        subject: "Profitable Project 1 comms design",
        state: "draft",
        issue_date: "2025-08-29",
        due_date: "2025-09-28",
        currency: "GBP",
        client: { id: 20030, name: "Acme Corp" },
        project_id: 100030,
      },
    ];
    const invoicesObj = {};
    invoicesArray.forEach(invoice => {
        invoicesObj[invoice.id] = invoice;
    });

  return invoicesArray;
}