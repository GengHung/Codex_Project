const STORAGE_KEY = "minimal-todo-tasks";
const VIEW_MODE_KEY = "minimal-todo-view-mode";

const form = document.querySelector("#task-form");
const input = document.querySelector("#task-input");
const assigneeInput = document.querySelector("#assignee-input");
const categorySelect = document.querySelector("#category-select");
const taskList = document.querySelector("#task-list");
const emptyState = document.querySelector("#empty-state");
const summary = document.querySelector("#task-summary");
const filterButtons = document.querySelectorAll(".filter");
const exportButton = document.querySelector("#export-button");
const submitButton = document.querySelector("#submit-button");
const cancelButton = document.querySelector("#cancel-button");
const viewButtons = document.querySelectorAll(".view-button");

let tasks = loadTasks();
let activeFilter = "all";
let editingTaskId = null;
let viewMode = loadViewMode();

function loadTasks() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    return Array.isArray(saved) ? saved : [];
  } catch {
    return [];
  }
}

function saveTasks() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(tasks));
}

function loadViewMode() {
  const savedMode = localStorage.getItem(VIEW_MODE_KEY);
  return ["list", "card", "board"].includes(savedMode) ? savedMode : "list";
}

function getLocalDate(value = new Date()) {
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getTaskDateKey(task) {
  if (!task.createdAt) return "unknown";
  const date = new Date(task.createdAt);
  return Number.isNaN(date.getTime()) ? "unknown" : getLocalDate(date);
}

function formatDateLabel(dateKey) {
  if (dateKey === "unknown") return "日期未記錄";

  const today = getLocalDate();
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  if (dateKey === today) return "今天";
  if (dateKey === getLocalDate(yesterday)) return "昨天";

  const [year, month, day] = dateKey.split("-").map(Number);
  return new Date(year, month - 1, day).toLocaleDateString("zh-TW", {
    year: "numeric",
    month: "long",
    day: "numeric",
    weekday: "short",
  });
}

function escapeXml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function exportTodayTasks() {
  const today = getLocalDate();
  const todayTasks = tasks.filter((task) => {
    if (!task.createdAt) return true;
    return getLocalDate(new Date(task.createdAt)) === today;
  });

  if (!todayTasks.length) {
    window.alert("今天尚未新增任何任務。");
    return;
  }

  const rows = todayTasks.map((task) => `
    <Row>
      <Cell><Data ss:Type="String">${escapeXml(task.text)}</Data></Cell>
      <Cell><Data ss:Type="String">${escapeXml(task.assignee || "未指派")}</Data></Cell>
      <Cell><Data ss:Type="String">${task.category === "work" ? "工作" : "生活"}</Data></Cell>
      <Cell><Data ss:Type="String">${task.completed ? "已完成" : "未完成"}</Data></Cell>
      <Cell><Data ss:Type="String">${escapeXml(task.createdAt ? new Date(task.createdAt).toLocaleString("zh-TW") : today)}</Data></Cell>
    </Row>`).join("");

  const workbook = `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Worksheet ss:Name="今日待辦">
    <Table>
      <Row>
        <Cell><Data ss:Type="String">任務</Data></Cell>
        <Cell><Data ss:Type="String">負責人</Data></Cell>
        <Cell><Data ss:Type="String">分類</Data></Cell>
        <Cell><Data ss:Type="String">狀態</Data></Cell>
        <Cell><Data ss:Type="String">新增時間</Data></Cell>
      </Row>${rows}
    </Table>
  </Worksheet>
</Workbook>`;

  const blob = new Blob([workbook], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `待辦清單_${today}.xls`;
  link.click();
  URL.revokeObjectURL(url);
}

function getTaskStatus(task) {
  if (["todo", "process", "done"].includes(task.status)) return task.status;
  return task.completed ? "done" : "todo";
}

function createTaskElement(task, draggable = false) {
  const item = document.createElement("li");
  item.className = `task-item${task.completed ? " completed" : ""}${task.id === editingTaskId ? " editing" : ""}`;
  item.dataset.id = task.id;
  item.draggable = draggable;

  const checkButton = document.createElement("button");
  checkButton.className = "task-item__check";
  checkButton.type = "button";
  checkButton.dataset.action = "toggle";
  checkButton.setAttribute("aria-label", task.completed ? "標記為未完成" : "標記為完成");
  checkButton.textContent = "✓";

  const content = document.createElement("div");
  content.className = "task-item__content";
  content.tabIndex = 0;
  content.setAttribute("role", "button");
  content.setAttribute("aria-label", `修改「${task.text}」`);

  const text = document.createElement("p");
  text.className = "task-item__text";
  text.textContent = task.text;

  const badge = document.createElement("span");
  badge.className = `badge badge--${task.category}`;
  badge.textContent = task.category === "work" ? "工作" : "生活";

  const assignee = document.createElement("span");
  assignee.className = "assignee";
  assignee.textContent = `負責人：${task.assignee || "未指派"}`;

  const meta = document.createElement("div");
  meta.className = "task-item__meta";
  meta.append(badge, assignee);

  const deleteButton = document.createElement("button");
  deleteButton.className = "task-item__delete";
  deleteButton.type = "button";
  deleteButton.dataset.action = "delete";
  deleteButton.setAttribute("aria-label", `刪除「${task.text}」`);
  deleteButton.textContent = "×";

  content.append(text, meta);
  item.append(checkButton, content, deleteButton);
  return item;
}

function createDateSection(dateKey, groupedTasks) {
  const section = document.createElement("section");
  section.className = "date-section";

  const header = document.createElement("div");
  header.className = "date-section__header";

  const title = document.createElement("h2");
  title.className = "date-section__title";
  title.textContent = formatDateLabel(dateKey);

  const count = document.createElement("span");
  count.className = "date-section__count";
  count.textContent = `${groupedTasks.length} 項`;

  const list = document.createElement("ul");
  list.className = "date-section__tasks";
  list.append(...groupedTasks.map(createTaskElement));

  header.append(title, count);
  section.append(header, list);
  return section;
}

function createKanbanBoard(visibleTasks) {
  const columns = [
    { status: "todo", label: "To-do" },
    { status: "process", label: "Process" },
    { status: "done", label: "Done" },
  ];

  const board = document.createElement("div");
  board.className = "kanban-board";

  columns.forEach(({ status, label }) => {
    const columnTasks = visibleTasks.filter((task) => getTaskStatus(task) === status);
    const column = document.createElement("section");
    column.className = "kanban-column";
    column.dataset.status = status;

    const header = document.createElement("div");
    header.className = "kanban-column__header";

    const title = document.createElement("h2");
    title.className = "kanban-column__title";
    title.textContent = label;

    const count = document.createElement("span");
    count.className = "kanban-column__count";
    count.textContent = columnTasks.length;
    count.setAttribute("aria-label", `${columnTasks.length} 項任務`);

    const list = document.createElement("ul");
    list.className = "kanban-column__tasks";
    list.append(...columnTasks.map((task) => createTaskElement(task, true)));

    header.append(title, count);
    column.append(header, list);
    board.append(column);
  });

  return board;
}

function render() {
  const visibleTasks = activeFilter === "all"
    ? tasks
    : tasks.filter((task) => task.category === activeFilter);

  taskList.className = `task-list view--${viewMode}`;
  if (viewMode === "board") {
    taskList.replaceChildren(createKanbanBoard(visibleTasks));
  } else {
    const groups = new Map();
    visibleTasks.forEach((task) => {
      const dateKey = getTaskDateKey(task);
      if (!groups.has(dateKey)) groups.set(dateKey, []);
      groups.get(dateKey).push(task);
    });

    const dateKeys = [...groups.keys()].sort((a, b) => {
      if (a === "unknown") return 1;
      if (b === "unknown") return -1;
      return b.localeCompare(a);
    });
    taskList.replaceChildren(...dateKeys.map((dateKey) => createDateSection(dateKey, groups.get(dateKey))));
  }
  emptyState.hidden = visibleTasks.length > 0;

  viewButtons.forEach((button) => {
    const isActive = button.dataset.view === viewMode;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  });

  const remaining = tasks.filter((task) => !task.completed).length;
  summary.textContent = tasks.length
    ? `還有 ${remaining} 項任務待完成`
    : "今天要完成什麼？";
}

function enterEditMode(taskId) {
  const task = tasks.find((item) => item.id === taskId);
  if (!task) return;

  editingTaskId = task.id;
  input.value = task.text;
  assigneeInput.value = task.assignee || "";
  categorySelect.value = task.category;
  submitButton.textContent = "修改";
  cancelButton.hidden = false;
  render();
  input.focus();
}

function exitEditMode() {
  editingTaskId = null;
  form.reset();
  submitButton.textContent = "新增";
  cancelButton.hidden = true;
  render();
  input.focus();
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  const text = input.value.trim();
  const assignee = assigneeInput.value.trim();
  if (!text || !assignee) return;

  if (editingTaskId) {
    tasks = tasks.map((task) => task.id === editingTaskId
      ? { ...task, text, assignee, category: categorySelect.value }
      : task);
  } else {
    tasks.unshift({
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`,
      text,
      assignee,
      category: categorySelect.value,
      completed: false,
      status: "todo",
      createdAt: new Date().toISOString(),
    });
  }

  saveTasks();
  exitEditMode();
});

taskList.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-action]");
  const item = event.target.closest(".task-item");
  if (!item) return;

  if (button?.dataset.action === "toggle") {
    tasks = tasks.map((task) => task.id === item.dataset.id
      ? {
          ...task,
          completed: !task.completed,
          status: task.completed ? "todo" : "done",
        }
      : task);
    saveTasks();
    render();
  } else if (button?.dataset.action === "delete") {
    tasks = tasks.filter((task) => task.id !== item.dataset.id);
    saveTasks();
    if (editingTaskId === item.dataset.id) {
      exitEditMode();
    } else {
      render();
    }
  } else if (event.target.closest(".task-item__content")) {
    enterEditMode(item.dataset.id);
  }
});

taskList.addEventListener("keydown", (event) => {
  const content = event.target.closest(".task-item__content");
  if (!content || (event.key !== "Enter" && event.key !== " ")) return;

  event.preventDefault();
  const item = content.closest(".task-item");
  enterEditMode(item.dataset.id);
});

taskList.addEventListener("dragstart", (event) => {
  const item = event.target.closest(".task-item[draggable='true']");
  if (!item) return;
  event.dataTransfer.effectAllowed = "move";
  event.dataTransfer.setData("text/plain", item.dataset.id);
  item.classList.add("dragging");
});

taskList.addEventListener("dragend", (event) => {
  event.target.closest(".task-item")?.classList.remove("dragging");
  taskList.querySelectorAll(".kanban-column").forEach((column) => column.classList.remove("drag-over"));
});

taskList.addEventListener("dragover", (event) => {
  const column = event.target.closest(".kanban-column");
  if (!column) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = "move";
  taskList.querySelectorAll(".kanban-column").forEach((item) => item.classList.toggle("drag-over", item === column));
});

taskList.addEventListener("dragleave", (event) => {
  const column = event.target.closest(".kanban-column");
  if (column && !column.contains(event.relatedTarget)) column.classList.remove("drag-over");
});

taskList.addEventListener("drop", (event) => {
  const column = event.target.closest(".kanban-column");
  if (!column) return;
  event.preventDefault();

  const taskId = event.dataTransfer.getData("text/plain");
  const status = column.dataset.status;
  tasks = tasks.map((task) => task.id === taskId
    ? { ...task, status, completed: status === "done" }
    : task);
  saveTasks();
  render();
});

filterButtons.forEach((button) => {
  button.addEventListener("click", () => {
    activeFilter = button.dataset.filter;
    filterButtons.forEach((filter) => filter.classList.toggle("active", filter === button));
    render();
  });
});

viewButtons.forEach((button) => {
  button.addEventListener("click", () => {
    viewMode = button.dataset.view;
    localStorage.setItem(VIEW_MODE_KEY, viewMode);
    render();
  });
});

exportButton.addEventListener("click", exportTodayTasks);
cancelButton.addEventListener("click", exitEditMode);

render();
