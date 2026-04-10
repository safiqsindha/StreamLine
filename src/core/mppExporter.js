/**
 * Streamline MS Project Exporter
 * Generates an MS Project XML file (.xml) that can be opened directly in MS Project.
 * This enables write-back / sync from Streamline -> MS Project.
 *
 * The XML schema follows the Microsoft Project 2007+ format.
 */

const { TaskType, DepType } = require("./dataModel");

// Dependency type code: MS Project uses 0=FF, 1=FS, 2=SF, 3=SS
const DEP_TYPE_CODE = {
  FF: 0,
  FS: 1,
  SF: 2,
  SS: 3,
};

/**
 * Export tasks to MS Project XML format.
 * @param {Array} tasks - Task objects from dataModel
 * @param {Array} swimLanes - Swim lane objects (for hierarchy)
 * @param {string} projectName
 * @returns {string} XML content
 */
function exportToMppXml(tasks, swimLanes, projectName = "Streamline Export") {
  const now = new Date().toISOString();
  const taskNameToUid = new Map();

  // Assign UIDs starting at 1 (UID 0 is reserved for project summary)
  let nextUid = 1;
  const summaryLookup = new Map(); // swimLane name -> summary UID

  // Build header
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
  xml += `<Project xmlns="http://schemas.microsoft.com/project">\n`;
  xml += `  <SaveVersion>14</SaveVersion>\n`;
  xml += `  <Name>${esc(projectName)}</Name>\n`;
  xml += `  <Title>${esc(projectName)}</Title>\n`;
  xml += `  <Author>Streamline</Author>\n`;
  xml += `  <CreationDate>${now}</CreationDate>\n`;
  xml += `  <LastSaved>${now}</LastSaved>\n`;
  xml += `  <ScheduleFromStart>1</ScheduleFromStart>\n`;
  xml += `  <StartDate>${findEarliestDate(tasks)}</StartDate>\n`;
  xml += `  <FinishDate>${findLatestDate(tasks)}</FinishDate>\n`;
  xml += `  <FYStartDate>1</FYStartDate>\n`;
  xml += `  <CriticalSlackLimit>0</CriticalSlackLimit>\n`;
  xml += `  <CurrencyDigits>2</CurrencyDigits>\n`;
  xml += `  <CurrencySymbol>$</CurrencySymbol>\n`;
  xml += `  <CurrencySymbolPosition>0</CurrencySymbolPosition>\n`;
  xml += `  <CalendarUID>1</CalendarUID>\n`;
  xml += `  <DefaultStartTime>08:00:00</DefaultStartTime>\n`;
  xml += `  <DefaultFinishTime>17:00:00</DefaultFinishTime>\n`;
  xml += `  <MinutesPerDay>480</MinutesPerDay>\n`;
  xml += `  <MinutesPerWeek>2400</MinutesPerWeek>\n`;
  xml += `  <DaysPerMonth>20</DaysPerMonth>\n`;
  xml += `  <DefaultTaskType>0</DefaultTaskType>\n`;
  xml += `  <DefaultFixedCostAccrual>3</DefaultFixedCostAccrual>\n`;
  xml += `  <DefaultStandardRate>0</DefaultStandardRate>\n`;
  xml += `  <DefaultOvertimeRate>0</DefaultOvertimeRate>\n`;
  xml += `  <DurationFormat>7</DurationFormat>\n`;
  xml += `  <WorkFormat>2</WorkFormat>\n`;
  xml += `  <EditableActualCosts>0</EditableActualCosts>\n`;
  xml += `  <HonorConstraints>0</HonorConstraints>\n`;
  xml += `  <InsertedProjectsLikeSummary>1</InsertedProjectsLikeSummary>\n`;
  xml += `  <MultipleCriticalPaths>0</MultipleCriticalPaths>\n`;
  xml += `  <NewTasksEffortDriven>0</NewTasksEffortDriven>\n`;
  xml += `  <NewTasksEstimated>1</NewTasksEstimated>\n`;
  xml += `  <SplitsInProgressTasks>0</SplitsInProgressTasks>\n`;
  xml += `  <SpreadActualCost>0</SpreadActualCost>\n`;
  xml += `  <SpreadPercentComplete>0</SpreadPercentComplete>\n`;
  xml += `  <TaskUpdatesResource>1</TaskUpdatesResource>\n`;
  xml += `  <FiscalYearStart>0</FiscalYearStart>\n`;
  xml += `  <WeekStartDay>1</WeekStartDay>\n`;
  xml += `  <MoveCompletedEndsBack>0</MoveCompletedEndsBack>\n`;
  xml += `  <MoveRemainingStartsBack>0</MoveRemainingStartsBack>\n`;
  xml += `  <MoveRemainingStartsForward>0</MoveRemainingStartsForward>\n`;
  xml += `  <MoveCompletedEndsForward>0</MoveCompletedEndsForward>\n`;
  xml += `  <BaselineForEarnedValue>0</BaselineForEarnedValue>\n`;
  xml += `  <AutoAddResources>1</AutoAddResources>\n`;
  xml += `  <StatusDate>${now}</StatusDate>\n`;
  xml += `  <CurrentDate>${now}</CurrentDate>\n`;
  xml += `  <MicrosoftProjectServerURL>0</MicrosoftProjectServerURL>\n`;
  xml += `  <Autolink>1</Autolink>\n`;
  xml += `  <NewTaskStartDate>0</NewTaskStartDate>\n`;
  xml += `  <DefaultTaskEVMethod>0</DefaultTaskEVMethod>\n`;
  xml += `  <ProjectExternallyEdited>0</ProjectExternallyEdited>\n`;
  xml += `  <ExtendedCreationDate>${now}</ExtendedCreationDate>\n`;
  xml += `  <ActualsInSync>0</ActualsInSync>\n`;
  xml += `  <RemoveFileProperties>0</RemoveFileProperties>\n`;
  xml += `  <AdminProject>0</AdminProject>\n`;

  // Calendars (simple: one Standard)
  xml += `  <Calendars>\n`;
  xml += `    <Calendar>\n`;
  xml += `      <UID>1</UID>\n`;
  xml += `      <Name>Standard</Name>\n`;
  xml += `      <IsBaseCalendar>1</IsBaseCalendar>\n`;
  xml += `      <BaseCalendarUID>-1</BaseCalendarUID>\n`;
  xml += `      <WeekDays>\n`;
  // Sun=1, Mon=2, ..., Sat=7 in MS Project
  for (let d = 1; d <= 7; d++) {
    const isWorking = d >= 2 && d <= 6 ? 1 : 0;
    xml += `        <WeekDay>\n`;
    xml += `          <DayType>${d}</DayType>\n`;
    xml += `          <DayWorking>${isWorking}</DayWorking>\n`;
    if (isWorking) {
      xml += `          <WorkingTimes>\n`;
      xml += `            <WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime>\n`;
      xml += `            <WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime>\n`;
      xml += `          </WorkingTimes>\n`;
    }
    xml += `        </WeekDay>\n`;
  }
  xml += `      </WeekDays>\n`;
  xml += `    </Calendar>\n`;
  xml += `  </Calendars>\n`;

  // Tasks
  xml += `  <Tasks>\n`;

  // Project summary task (UID 0)
  xml += `    <Task>\n`;
  xml += `      <UID>0</UID>\n`;
  xml += `      <ID>0</ID>\n`;
  xml += `      <Name>${esc(projectName)}</Name>\n`;
  xml += `      <Type>1</Type>\n`;
  xml += `      <IsNull>0</IsNull>\n`;
  xml += `      <Summary>1</Summary>\n`;
  xml += `      <OutlineNumber>0</OutlineNumber>\n`;
  xml += `      <OutlineLevel>0</OutlineLevel>\n`;
  xml += `      <Priority>500</Priority>\n`;
  xml += `      <Start>${findEarliestDate(tasks)}</Start>\n`;
  xml += `      <Finish>${findLatestDate(tasks)}</Finish>\n`;
  xml += `      <Milestone>0</Milestone>\n`;
  xml += `    </Task>\n`;

  // Swim lane summary tasks + child tasks
  let taskId = 1;
  for (const lane of swimLanes) {
    const summaryUid = nextUid++;
    summaryLookup.set(lane.name, summaryUid);

    xml += taskXml({
      uid: summaryUid,
      id: taskId++,
      name: lane.name,
      outlineLevel: 1,
      isSummary: true,
      isMilestone: false,
      start: findEarliestDate(lane.tasks),
      finish: findLatestDate(lane.tasks),
    });

    // Top-level tasks in this swim lane
    const topTasks = lane.topLevelTasks || lane.tasks;
    for (const task of topTasks) {
      const uid = nextUid++;
      taskNameToUid.set(task.name, uid);
      xml += taskXml({
        uid,
        id: taskId++,
        name: task.name,
        outlineLevel: 2,
        isSummary: false,
        isMilestone: task.type === TaskType.MILESTONE,
        start: dateStr(task.startDate),
        finish: dateStr(task.endDate || task.startDate),
        percentComplete: task.percentComplete,
        baselineStart: dateStr(task.plannedStartDate),
        baselineFinish: dateStr(task.plannedEndDate),
      });
    }

    // Sub-lane tasks
    if (lane.subLanes) {
      for (const subLane of lane.subLanes) {
        const subSummaryUid = nextUid++;
        xml += taskXml({
          uid: subSummaryUid,
          id: taskId++,
          name: subLane.name,
          outlineLevel: 2,
          isSummary: true,
          isMilestone: false,
          start: findEarliestDate(subLane.tasks),
          finish: findLatestDate(subLane.tasks),
        });

        for (const task of subLane.tasks) {
          const uid = nextUid++;
          taskNameToUid.set(task.name, uid);
          xml += taskXml({
            uid,
            id: taskId++,
            name: task.name,
            outlineLevel: 3,
            isSummary: false,
            isMilestone: task.type === TaskType.MILESTONE,
            start: dateStr(task.startDate),
            finish: dateStr(task.endDate || task.startDate),
            percentComplete: task.percentComplete,
            baselineStart: dateStr(task.plannedStartDate),
            baselineFinish: dateStr(task.plannedEndDate),
          });
        }
      }
    }
  }

  // Add predecessor links to each task
  // We need to regenerate task XML since predecessors reference UIDs by task
  // For simplicity, we append a separate pass using a string replacement approach
  // Actually, let's just include it in taskXml — rewrite above... too late, do post-processing

  xml += `  </Tasks>\n`;
  xml += `  <Resources></Resources>\n`;
  xml += `  <Assignments></Assignments>\n`;
  xml += `</Project>\n`;

  // Post-process: inject predecessor links
  for (const task of tasks) {
    const taskUid = taskNameToUid.get(task.name);
    if (!taskUid || !task.dependencies || task.dependencies.length === 0) continue;

    let predXml = "";
    for (const depId of task.dependencies) {
      const depTask = tasks.find((t) => t.id === depId);
      if (!depTask) continue;
      const predUid = taskNameToUid.get(depTask.name);
      if (!predUid) continue;

      const depInfo = task.dependencyTypes.get(depId) || { type: DepType.FS, lagDays: 0 };
      const typeCode = DEP_TYPE_CODE[depInfo.type] || 1;
      // MS Project uses tenths of minutes for lag
      const lagMinutes = depInfo.lagDays * 480;

      predXml += `      <PredecessorLink>\n`;
      predXml += `        <PredecessorUID>${predUid}</PredecessorUID>\n`;
      predXml += `        <Type>${typeCode}</Type>\n`;
      predXml += `        <CrossProject>0</CrossProject>\n`;
      predXml += `        <LinkLag>${lagMinutes * 10}</LinkLag>\n`;
      predXml += `        <LagFormat>7</LagFormat>\n`;
      predXml += `      </PredecessorLink>\n`;
    }

    // Inject into task XML (find the closing </Task> for this UID and insert before)
    const marker = `      <UID>${taskUid}</UID>\n`;
    const start = xml.indexOf(marker);
    if (start === -1) continue;
    const endMarker = `    </Task>\n`;
    const end = xml.indexOf(endMarker, start);
    if (end === -1) continue;

    xml = xml.slice(0, end) + predXml + xml.slice(end);
  }

  return xml;
}

/**
 * Render a single task in MS Project XML format.
 */
function taskXml(opts) {
  const duration = durationFmt(opts.start, opts.finish, opts.isMilestone);
  let xml = `    <Task>\n`;
  xml += `      <UID>${opts.uid}</UID>\n`;
  xml += `      <ID>${opts.id}</ID>\n`;
  xml += `      <Name>${esc(opts.name)}</Name>\n`;
  xml += `      <Type>0</Type>\n`;
  xml += `      <IsNull>0</IsNull>\n`;
  xml += `      <CreateDate>${new Date().toISOString()}</CreateDate>\n`;
  xml += `      <WBS>${opts.id}</WBS>\n`;
  xml += `      <OutlineNumber>${opts.id}</OutlineNumber>\n`;
  xml += `      <OutlineLevel>${opts.outlineLevel}</OutlineLevel>\n`;
  xml += `      <Priority>500</Priority>\n`;
  xml += `      <Start>${opts.start}</Start>\n`;
  xml += `      <Finish>${opts.finish}</Finish>\n`;
  xml += `      <Duration>${duration}</Duration>\n`;
  xml += `      <DurationFormat>7</DurationFormat>\n`;
  xml += `      <Work>${duration}</Work>\n`;
  xml += `      <Summary>${opts.isSummary ? 1 : 0}</Summary>\n`;
  xml += `      <Milestone>${opts.isMilestone ? 1 : 0}</Milestone>\n`;
  xml += `      <Active>1</Active>\n`;
  xml += `      <Manual>0</Manual>\n`;
  xml += `      <ManualStart>${opts.start}</ManualStart>\n`;
  xml += `      <ManualFinish>${opts.finish}</ManualFinish>\n`;
  xml += `      <ManualDuration>${duration}</ManualDuration>\n`;
  if (opts.percentComplete !== undefined && opts.percentComplete !== null) {
    xml += `      <PercentComplete>${Math.round(opts.percentComplete)}</PercentComplete>\n`;
  } else {
    xml += `      <PercentComplete>0</PercentComplete>\n`;
  }
  if (opts.baselineStart && opts.baselineFinish) {
    xml += `      <Baseline>\n`;
    xml += `        <Number>0</Number>\n`;
    xml += `        <Start>${opts.baselineStart}</Start>\n`;
    xml += `        <Finish>${opts.baselineFinish}</Finish>\n`;
    xml += `        <Duration>${durationFmt(opts.baselineStart, opts.baselineFinish, false)}</Duration>\n`;
    xml += `        <DurationFormat>7</DurationFormat>\n`;
    xml += `        <Work>${durationFmt(opts.baselineStart, opts.baselineFinish, false)}</Work>\n`;
    xml += `      </Baseline>\n`;
  }
  xml += `      <FixedCostAccrual>3</FixedCostAccrual>\n`;
  xml += `      <ConstraintType>0</ConstraintType>\n`;
  xml += `      <CalendarUID>-1</CalendarUID>\n`;
  xml += `      <EffortDriven>0</EffortDriven>\n`;
  xml += `      <Estimated>1</Estimated>\n`;
  xml += `      <IgnoreResourceCalendar>0</IgnoreResourceCalendar>\n`;
  xml += `      <LevelAssignments>1</LevelAssignments>\n`;
  xml += `      <LevelingCanSplit>1</LevelingCanSplit>\n`;
  xml += `      <LevelingDelay>0</LevelingDelay>\n`;
  xml += `      <LevelingDelayFormat>8</LevelingDelayFormat>\n`;
  xml += `      <RegularWork>${duration}</RegularWork>\n`;
  xml += `    </Task>\n`;
  return xml;
}

function dateStr(date) {
  if (!date) return new Date().toISOString().slice(0, 19);
  const d = date instanceof Date ? date : new Date(date);
  return d.toISOString().slice(0, 19);
}

function durationFmt(start, finish, isMilestone) {
  if (isMilestone) return "PT0H0M0S";
  const s = new Date(start);
  const f = new Date(finish);
  const hours = Math.round((f - s) / (1000 * 60 * 60));
  return `PT${hours}H0M0S`;
}

function findEarliestDate(tasks) {
  let min = null;
  for (const t of tasks) {
    if (t.startDate && (!min || t.startDate < min)) min = t.startDate;
    if (t.plannedStartDate && (!min || t.plannedStartDate < min)) min = t.plannedStartDate;
  }
  return dateStr(min || new Date());
}

function findLatestDate(tasks) {
  let max = null;
  for (const t of tasks) {
    if (t.endDate && (!max || t.endDate > max)) max = t.endDate;
    if (t.startDate && (!max || t.startDate > max)) max = t.startDate;
  }
  return dateStr(max || new Date());
}

function esc(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Trigger browser download of an MS Project XML file.
 */
function downloadMppXml(tasks, swimLanes, projectName = "Streamline_Export") {
  const xml = exportToMppXml(tasks, swimLanes, projectName);
  const blob = new Blob([xml], { type: "application/xml" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${projectName.replace(/\s+/g, "_")}.xml`;
  a.click();
  URL.revokeObjectURL(url);
}

module.exports = { exportToMppXml, downloadMppXml };
