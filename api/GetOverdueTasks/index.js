const fetch = require('node-fetch');

// Main function handler
module.exports = async function (context, req) {
    const teamworkApiKey = process.env.TEAMWORK_API_KEY;
    const teamworkBaseUrl = process.env.TEAMWORK_BASE_URL;

    // Validate that environment variables are set
    if (!teamworkApiKey || !teamworkBaseUrl) {
        context.res = {
            status: 500,
            body: { error: "Server configuration error: API key or Base URL not set." }
        };
        return;
    }

    try {
        const allTasks = await fetchAllOverdueTasks(teamworkApiKey, teamworkBaseUrl);
        
        context.res = {
            status: 200,
            headers: { 'Content-Type': 'application/json' },
            body: allTasks
        };
    } catch (error) {
        context.log.error("Error fetching tasks:", error);
        context.res = {
            status: 500,
            body: { error: "Failed to fetch tasks from Teamwork API.", details: error.message }
        };
    }
};

// --- Helper function to fetch all pages of tasks ---
async function fetchAllOverdueTasks(apiKey, baseUrl) {
    let allOverdueTasks = [];
    let page = 1;
    const pageSize = 50;
    let hasMorePages = true;

    // Lookup tables for related data
    const projectsLookup = {};
    const tasklistsLookup = {};
    const usersLookup = {};

    const encodedCredentials = Buffer.from(`${apiKey}:X`).toString('base64');
    const headers = {
        "Authorization": `Basic ${encodedCredentials}`
    };

    while (hasMorePages) {
        const uri = `${baseUrl}/projects/api/v3/tasks.json?filter[completed]=false&filter[dueDate][lt]=now&page=${page}&pageSize=${pageSize}&include=projects,tasklists,users`;
        
        const response = await fetch(uri, { headers });
        if (!response.ok) {
            throw new Error(`Teamwork API request failed with status ${response.status}`);
        }
        const data = await response.json();

        // Populate lookup tables from the 'included' data
        if (data.included) {
            if (data.included.users) {
                Object.values(data.included.users).forEach(user => usersLookup[user.id] = user);
            }
            if (data.included.projects) {
                Object.values(data.included.projects).forEach(project => projectsLookup[project.id] = project);
            }
            if (data.included.tasklists) {
                Object.values(data.included.tasklists).forEach(tasklist => tasklistsLookup[tasklist.id] = tasklist);
            }
        }

        if (data.tasks && data.tasks.length > 0) {
            const tasksWithDueDate = data.tasks.filter(task => task.dueDate);
            allOverdueTasks.push(...tasksWithDueDate);
        }

        // Check for more pages
        hasMorePages = data.meta && data.meta.page && data.meta.page.hasMore;
        page++;
    }

    // Process and format the final list
    return formatTasks(allOverdueTasks, projectsLookup, tasklistsLookup, usersLookup, baseUrl);
}

// --- Helper function to format the final output ---
function formatTasks(tasks, projects, tasklists, users, baseUrl) {
    return tasks.map(task => {
        // Find project name
        const tasklist = tasklists[task.tasklistId];
        const project = tasklist ? projects[tasklist.projectId] : null;
        const projectName = project ? project.name : "Project Not Found";

        // Find assignee names
        let assigneeString = "Unassigned";
        if (task.assigneeUserIds && task.assigneeUserIds.length > 0) {
            const assigneeNames = task.assigneeUserIds.map(id => {
                const user = users[id];
                return user ? `${user.firstName} ${user.lastName}` : "Unknown User";
            });
            assigneeString = assigneeNames.join(', ');
        }

        return {
            TaskID: task.id,
            Project: projectName,
            Task: task.name,
            AssignedTo: assigneeString,
            DueDate: task.dueDate,
            URL: `${baseUrl}/app/tasks/${task.id}`
        };
    }).sort((a, b) => new Date(a.DueDate) - new Date(b.DueDate)); // Sort by due date
}
