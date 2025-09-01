function getDummyTimeEntries() {
  const projects = getDummyProjects().projects;
  const users = getDummyUsers().users;

  // Get current financial year start and end
  const currentDate = new Date();
  const currentYear = currentDate.getFullYear();
  const currentMonth = currentDate.getMonth();
  const currentDay = currentDate.getDate();
  const startYear = currentMonth > 3 || (currentMonth === 3 && currentDay >= 6) ? currentYear : currentYear - 1;
  const endYear = startYear + 1;
  const fyStart = new Date(startYear, 3, 1); // April 1st
  const fyEnd = new Date(endYear, 3, 1); // April 1st next year

  function randomDateInFY() {
    const start = fyStart.getTime();
    const end = fyEnd.getTime();
    return new Date(start + Math.random() * (end - start));
  }

  let entryId = 1000000000;

  const time_entries = projects.flatMap(project =>
    users.map(user => {
      const spentDateObj = randomDateInFY();
      const spent_date = spentDateObj.toISOString().slice(0, 10);
      const created_at = spentDateObj.toISOString();
      const updated_at = spentDateObj.toISOString();

      return {
        id: entryId++,
        spent_date,
        hours: Math.round((Math.random() * 7 + 1) * 100) / 100,
        hours_without_timer: null,
        rounded_hours: null,
        notes: `Dummy entry for ${user.name} on ${project.name}`,
        is_locked: false,
        locked_reason: null,
        approval_status: "unsubmitted",
        is_closed: false,
        is_billed: false,
        timer_started_at: null,
        started_time: null,
        ended_time: null,
        is_running: false,
        billable: true,
        budgeted: true,
        billable_rate: project.hourly_rate || 100,
        cost_rate: null,
        created_at,
        updated_at,
        user: {
          id: user.id,
          name: user.name
        },
        client: {
          id: project.client.id,
          name: project.client.name,
          currency: "GBP"
        },
        project: {
          id: project.id,
          name: project.name,
          code: project.code
        },
        task: {
          id: 1000 + (entryId % 100),
          name: "Dummy Task"
        },
        user_assignment: {
          id: 500000000 + user.id,
          is_project_manager: false,
          is_active: true,
          use_default_rates: true,
          budget: null,
          created_at,
          updated_at,
          hourly_rate: project.hourly_rate || 100
        },
        task_assignment: {
          id: 600000000 + user.id,
          billable: true,
          is_active: true,
          created_at,
          updated_at,
          hourly_rate: project.hourly_rate || 100,
          budget: null
        },
        invoice: null,
        external_reference: null
      };
    })
  );

  return { time_entries };
}