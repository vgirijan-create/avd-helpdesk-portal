// MSAL configuration - replace with your own values
const msalConfig = {
    auth: {
        clientId: "YOUR_CLIENT_ID", // Replace with your app registration client ID
        authority: "https://login.microsoftonline.com/YOUR_TENANT_ID", // Replace with your tenant ID
        redirectUri: window.location.origin, // This should match what's registered in Azure AD
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

// Azure API endpoints
const apiConfig = {
    subscriptionsEndpoint: "https://management.azure.com/subscriptions?api-version=2020-01-01",
    resourceGroupsEndpoint: "https://management.azure.com/subscriptions/{subscriptionId}/resourcegroups?api-version=2021-04-01",
    hostPoolsEndpoint: "https://management.azure.com/subscriptions/{subscriptionId}/resourceGroups/{resourceGroup}/providers/Microsoft.DesktopVirtualization/hostPools?api-version=2022-02-10-preview",
    sessionHostsEndpoint: "https://management.azure.com/subscriptions/{subscriptionId}/resourceGroups/{resourceGroup}/providers/Microsoft.DesktopVirtualization/hostPools/{hostPool}/sessionHosts?api-version=2022-02-10-preview",
    userSessionsEndpoint: "https://management.azure.com/subscriptions/{subscriptionId}/resourceGroups/{resourceGroup}/providers/Microsoft.DesktopVirtualization/hostPools/{hostPool}/sessionHosts/{sessionHost}/userSessions?api-version=2022-02-10-preview",
    logoffUserEndpoint: "https://management.azure.com/subscriptions/{subscriptionId}/resourceGroups/{resourceGroup}/providers/Microsoft.DesktopVirtualization/hostPools/{hostPool}/sessionHosts/{sessionHost}/userSessions/{sessionId}?api-version=2022-02-10-preview"
};

// Login request with required scopes for Azure Management API
const loginRequest = {
    scopes: ["https://management.azure.com/user_impersonation"]
};

// Initialize MSAL instance
const msalInstance = new msal.PublicClientApplication(msalConfig);

// DOM elements
const loginBtn = document.getElementById('loginBtn');
const logoutBtn = document.getElementById('logoutBtn');
const loadSessionsBtn = document.getElementById('loadSessionsBtn');
const subscriptionSelect = document.getElementById('subscriptionSelect');
const resourceGroupSelect = document.getElementById('resourceGroupSelect');
const hostPoolSelect = document.getElementById('hostPoolSelect');
const sessionsContainer = document.getElementById('sessionsContainer');
const loadingSpinner = document.getElementById('loadingSpinner');
const logoffModal = new bootstrap.Modal(document.getElementById('logoffModal'));
const userToLogOff = document.getElementById('userToLogOff');
const confirmLogoffBtn = document.getElementById('confirmLogoffBtn');

// Keep track of session data for logoff operations
let currentSessionData = {};

// Check if user is already logged in
window.addEventListener('load', async () => {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        // User is signed in
        loginBtn.style.display = 'none';
        logoutBtn.style.display = 'inline-block';
        await loadSubscriptions();
    }
});

// Handle login
loginBtn.addEventListener('click', async () => {
    try {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        console.log('Login successful', loginResponse);
        loginBtn.style.display = 'none';
        logoutBtn.style.display = 'inline-block';
        await loadSubscriptions();
    } catch (error) {
        console.error('Login failed', error);
        alert('Login failed: ' + error.message);
    }
});

// Handle logout
logoutBtn.addEventListener('click', () => {
    msalInstance.logout();
});

// Load subscriptions after login
async function loadSubscriptions() {
    showLoading(true);
    try {
        const token = await getToken();
        const response = await fetch(apiConfig.subscriptionsEndpoint, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            populateSelect(subscriptionSelect, data.value.map(sub => ({
                id: sub.subscriptionId,
                name: sub.displayName
            })));
            
            subscriptionSelect.disabled = false;
        } else {
            throw new Error(`Failed to load subscriptions: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error('Error loading subscriptions:', error);
        alert('Failed to load subscriptions. Check console for details.');
    } finally {
        showLoading(false);
    }
}

// Load resource groups when subscription is selected
subscriptionSelect.addEventListener('change', async () => {
    if (subscriptionSelect.value === 'Select Subscription') {
        resourceGroupSelect.disabled = true;
        hostPoolSelect.disabled = true;
        loadSessionsBtn.disabled = true;
        return;
    }
    
    showLoading(true);
    resourceGroupSelect.disabled = true;
    hostPoolSelect.disabled = true;
    loadSessionsBtn.disabled = true;
    
    try {
        const token = await getToken();
        const subscriptionId = subscriptionSelect.value;
        const endpoint = apiConfig.resourceGroupsEndpoint.replace('{subscriptionId}', subscriptionId);
        
        const response = await fetch(endpoint, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            populateSelect(resourceGroupSelect, data.value.map(rg => ({
                id: rg.name,
                name: rg.name
            })));
            
            resourceGroupSelect.disabled = false;
        } else {
            throw new Error(`Failed to load resource groups: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error('Error loading resource groups:', error);
        alert('Failed to load resource groups. Check console for details.');
    } finally {
        showLoading(false);
    }
});

// Load host pools when resource group is selected
resourceGroupSelect.addEventListener('change', async () => {
    if (resourceGroupSelect.value === 'Select Resource Group') {
        hostPoolSelect.disabled = true;
        loadSessionsBtn.disabled = true;
        return;
    }
    
    showLoading(true);
    hostPoolSelect.disabled = true;
    loadSessionsBtn.disabled = true;
    
    try {
        const token = await getToken();
        const subscriptionId = subscriptionSelect.value;
        const resourceGroup = resourceGroupSelect.value;
        const endpoint = apiConfig.hostPoolsEndpoint
            .replace('{subscriptionId}', subscriptionId)
            .replace('{resourceGroup}', resourceGroup);
        
        const response = await fetch(endpoint, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            populateSelect(hostPoolSelect, data.value.map(hp => ({
                id: hp.name,
                name: hp.name
            })));
            
            hostPoolSelect.disabled = false;
            loadSessionsBtn.disabled = false;
        } else {
            throw new Error(`Failed to load host pools: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error('Error loading host pools:', error);
        alert('Failed to load host pools. Check console for details.');
    } finally {
        showLoading(false);
    }
});

// Load sessions when the button is clicked
loadSessionsBtn.addEventListener('click', async () => {
    const subscriptionId = subscriptionSelect.value;
    const resourceGroup = resourceGroupSelect.value;
    const hostPool = hostPoolSelect.value;
    
    if (!subscriptionId || !resourceGroup || !hostPool) {
        alert('Please select subscription, resource group, and host pool');
        return;
    }
    
    await loadSessionHosts(subscriptionId, resourceGroup, hostPool);
});

// Load session hosts and their user sessions
async function loadSessionHosts(subscriptionId, resourceGroup, hostPool) {
    showLoading(true);
    sessionsContainer.innerHTML = '';
    currentSessionData = {};
    
    try {
        const token = await getToken();
        const endpoint = apiConfig.sessionHostsEndpoint
            .replace('{subscriptionId}', subscriptionId)
            .replace('{resourceGroup}', resourceGroup)
            .replace('{hostPool}', hostPool);
        
        const response = await fetch(endpoint, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            
            if (data.value.length === 0) {
                sessionsContainer.innerHTML = '<div class="alert alert-info">No session hosts found in this host pool.</div>';
                return;
            }
            
            // Create a row for each session host
            for (const sessionHost of data.value) {
                const sessionHostName = sessionHost.name.split('/').pop();
                const hostCard = createSessionHostCard(sessionHostName);
                sessionsContainer.appendChild(hostCard);
                
                // Load user sessions for this host
                await loadUserSessions(subscriptionId, resourceGroup, hostPool, sessionHostName, hostCard);
            }
        } else {
            throw new Error(`Failed to load session hosts: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error('Error loading session hosts:', error);
        sessionsContainer.innerHTML = `<div class="alert alert-danger">Failed to load session hosts: ${error.message}</div>`;
    } finally {
        showLoading(false);
    }
}

// Load user sessions for a specific session host
async function loadUserSessions(subscriptionId, resourceGroup, hostPool, sessionHostName, hostCard) {
    try {
        const token = await getToken();
        const endpoint = apiConfig.userSessionsEndpoint
            .replace('{subscriptionId}', subscriptionId)
            .replace('{resourceGroup}', resourceGroup)
            .replace('{hostPool}', hostPool)
            .replace('{sessionHost}', sessionHostName);
        
        const response = await fetch(endpoint, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            const usersList = hostCard.querySelector('.users-list');
            
            if (data.value.length === 0) {
                usersList.innerHTML = '<p class="text-muted">No active user sessions</p>';
                return;
            }
            
            usersList.innerHTML = '';
            data.value.forEach(session => {
                const sessionId = session.name.split('/').pop();
                const userName = session.properties.userPrincipalName || 'Unknown User';
                const sessionState = session.properties.sessionState || 'Unknown';
                
                // Save session data for logoff operation
                currentSessionData[sessionId] = {
                    subscriptionId,
                    resourceGroup,
                    hostPool,
                    sessionHost: sessionHostName,
                    sessionId,
                    userName
                };
                
                const userItem = document.createElement('div');
                userItem.className = 'user-item p-2 border-bottom';
                userItem.innerHTML = `
                    <div class="d-flex justify-content-between align-items-center">
                        <div>
                            <strong>${userName}</strong>
                            <span class="badge ${sessionState === 'Active' ? 'bg-success' : 'bg-secondary'} ms-2">${sessionState}</span>
                        </div>
                        <button class="btn btn-sm btn-danger logoff-btn" data-session-id="${sessionId}">Log Off</button>
                    </div>
                `;
                usersList.appendChild(userItem);
                
                // Add click handler for logoff button
                const logoffBtn = userItem.querySelector('.logoff-btn');
                logoffBtn.addEventListener('click', () => {
                    prepareLogoffUser(sessionId);
                });
            });
        } else {
            throw new Error(`Failed to load user sessions: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error('Error loading user sessions:', error);
        const usersList = hostCard.querySelector('.users-list');
        usersList.innerHTML = `<div class="alert alert-danger">Failed to load user sessions: ${error.message}</div>`;
    }
}

// Create a card for a session host
function createSessionHostCard(sessionHostName) {
    const col = document.createElement('div');
    col.className = 'col-md-6 col-lg-4';
    
    col.innerHTML = `
        <div class="card session-card">
            <div class="card-header">
                <h5 class="card-title mb-0">${sessionHostName}</h5>
            </div>
            <div class="card-body">
                <h6>Active Sessions</h6>
                <div class="users-list">
                    <div class="spinner-border spinner-border-sm text-primary" role="status">
                        <span class="visually-hidden">Loading...</span>
                    </div>
                    <span class="ms-2">Loading sessions...</span>
                </div>
            </div>
        </div>
    `;
    
    return col;
}

// Prepare logoff modal
function prepareLogoffUser(sessionId) {
    const sessionData = currentSessionData[sessionId];
    if (!sessionData) {
        alert('Session data not found');
        return;
    }
    
    userToLogOff.textContent = sessionData.userName;
    
    // Set up confirmation button
    confirmLogoffBtn.onclick = async () => {
        logoffModal.hide();
        await logoffUser(sessionData);
    };
    
    logoffModal.show();
}

// Log off a user
async function logoffUser(sessionData) {
    showLoading(true);
    
    try {
        const token = await getToken();
        const endpoint = apiConfig.logoffUserEndpoint
            .replace('{subscriptionId}', sessionData.subscriptionId)
            .replace('{resourceGroup}', sessionData.resourceGroup)
            .replace('{hostPool}', sessionData.hostPool)
            .replace('{sessionHost}', sessionData.sessionHost)
            .replace('{sessionId}', sessionData.sessionId);
        
        const response = await fetch(endpoint, {
            method: 'DELETE',
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (response.ok) {
            alert(`Successfully logged off ${sessionData.userName}`);
            // Reload sessions to refresh the UI
            await loadSessionHosts(
                sessionData.subscriptionId,
                sessionData.resourceGroup,
                sessionData.hostPool
            );
        } else {
            throw new Error(`Failed to log off user: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error('Error logging off user:', error);
        alert(`Failed to log off user: ${error.message}`);
    } finally {
        showLoading(false);
    }
}

// Helper function to get access token
async function getToken() {
    try {
        const account = msalInstance.getAllAccounts()[0];
        const tokenResponse = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account
        });
        return tokenResponse.accessToken;
    } catch (error) {
        console.error('Error acquiring token silently', error);
        // Fall back to interactive method if silent acquisition fails
        const tokenResponse = await msalInstance.acquireTokenPopup(loginRequest);
        return tokenResponse.accessToken;
    }
}

// Helper function to populate select dropdowns
function populateSelect(selectElement, options) {
    // Clear existing options except the first one
    while (selectElement.options.length > 1) {
        selectElement.remove(1);
    }
    
    // Add new options
    options.forEach(option => {
        const optElement = document.createElement('option');
        optElement.value = option.id;
        optElement.textContent = option.name;
        selectElement.appendChild(optElement);
    });
}

// Helper function to show/hide loading spinner
function showLoading(show) {
    loadingSpinner.style.display = show ? 'block' : 'none';
}