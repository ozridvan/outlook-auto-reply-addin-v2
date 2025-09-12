let version = "1.0.10";

// Authentication configuration
const AUTH_CONFIG = {
    clientId: 'c2a8b650-50b2-446e-a8e9-bffa6698b77f', // This needs to be replaced with actual Azure App ID
    scopes: [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/MailboxSettings.ReadWrite',
        'https://graph.microsoft.com/People.Read'
    ]
};
// Office.js initialization
console.log('version: '+ version);

Office.onReady((info) => {
    console.log('Office.onReady called', info);
    if (info.host === Office.HostType.Outlook) {
        initializeApp();
        displayVersionInfo();
        checkApiSupport();
        checkOOFStatusNew();
    }
});

// Enhanced authentication manager
class AuthenticationManager {
    constructor() {
        this.accessToken = null;
        this.tokenExpiry = null;
        this.isSSO = false;
    }

    async getAccessToken(options = {}) {
        const defaultOptions = {
            allowSignInPrompt: true,
            allowConsentPrompt: true,
            forMSGraphAccess: true,
            ...options
        };

        try {
            // Check if we have a valid cached token
            if (this.accessToken && this.tokenExpiry && new Date() < this.tokenExpiry) {
                console.log('Using cached SSO token');
                return this.accessToken;
            }

            // Try SSO first
            console.log('Attempting SSO authentication...');
            const token = await Office.auth.getAccessToken(defaultOptions);
            
            // Cache the token (typically valid for 1 hour)
            this.accessToken = token;
            this.tokenExpiry = new Date(Date.now() + 50 * 60 * 1000); // 50 minutes
            this.isSSO = true;
            
            console.log('SSO authentication successful');
            return token;
            
        } catch (error) {
            console.error('SSO authentication failed:', error);
            
            // Handle specific error codes
            if (error.code === 13000) {
                throw new Error('SSO is not supported on this platform. Please use Outlook on the web or desktop.');
            } else if (error.code === 13001) {
                throw new Error('User is not signed in to Office.');
            } else if (error.code === 13002) {
                throw new Error('User consent is required.');
            } else if (error.code === 13003) {
                throw new Error('User consent was not granted.');
            } else if (error.code === 13006) {
                throw new Error('Current user is not in a supported Microsoft 365 subscription.');
            } else if (error.code === 13012) {
                throw new Error('Add-in is not configured for SSO. Please check manifest configuration.');
            }
            
            // Try fallback authentication methods
            return await this.tryFallbackAuth();
        }
    }

    async tryFallbackAuth() {
        console.log('Trying fallback authentication methods...');
        
        try {
            // Try using OfficeRuntime.auth if available (newer Office versions)
            if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.auth) {
                console.log('Trying OfficeRuntime.auth...');
                const token = await OfficeRuntime.auth.getAccessToken({
                    allowSignInPrompt: true,
                    allowConsentPrompt: true,
                    forMSGraphAccess: true
                });
                
                this.accessToken = token;
                this.tokenExpiry = new Date(Date.now() + 50 * 60 * 1000);
                this.isSSO = true;
                
                console.log('OfficeRuntime.auth successful');
                return token;
            }
        } catch (runtimeError) {
            console.error('OfficeRuntime.auth failed:', runtimeError);
        }
        
        // If all methods fail, throw an informative error
        throw new Error('Authentication failed. SSO is not supported on this platform. Please use Outlook on the web or a newer version of Outlook desktop.');
    }

    clearToken() {
        this.accessToken = null;
        this.tokenExpiry = null;
        this.isSSO = false;
    }

    isAuthenticated() {
        return this.accessToken && this.tokenExpiry && new Date() < this.tokenExpiry;
    }
}

// Global authentication manager instance
const authManager = new AuthenticationManager();

async function checkOOFStatusNew() {    
    try {
        console.log('Checking OOF status with new authentication...');
        
        const token = await authManager.getAccessToken({ allowSignInPrompt: false });
        
        const response = await fetch('https://graph.microsoft.com/v1.0/me/mailboxSettings/automaticRepliesSetting', {
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json'
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            console.log('OOF Status:', data);
            
            if (data.status === 'enabled' || data.status === 'scheduled') {
                showOofStatusBanner(data);
            }
        } else {
            console.error('Failed to get OOF status:', response.status, response.statusText);
        }
        
    } catch (error) {
        console.error('Error checking OOF status:', error);
        // Don't show error for initial check - it's not critical
        if (error.message.includes('not supported')) {
            console.log('OOF status check skipped - SSO not supported on this platform');
        }
    }
}

async function displayVersionInfo() {
    const versionInfoElement = document.getElementById("version-info");

    try {
        const clientInfo = getOutlookClientInfo();

        // Alınan bilgileri kullanıcıya göstermek için HTML içeriğini oluştur
        let htmlContent = `
            <ul>
                <li><strong>Uygulama:</strong> ${clientInfo.host}</li>
                <li><strong>Versiyon:</strong> ${clientInfo.applicationVersion}</li>
                <li><strong>Platform:</strong> ${clientInfo.platform}</li>
            </ul>
        `;
        versionInfoElement.innerHTML = htmlContent;

    } catch (error) {
        versionInfoElement.innerHTML = `<p style="color: red;">Hata: ${error.message}</p>`;
    }
}

/**
 * Office.context.diagnostics nesnesinden istemci bilgilerini alır ve
 * daha anlaşılır bir formatta bir nesne olarak döndürür.
 * @returns {object} { host, applicationVersion, platform }
 */
function getOutlookClientInfo() {
    // diagnostics API'sinin desteklenip desteklenmediğini kontrol et
    if (!Office.context.diagnostics) {
        throw new Error("Diagnostics API bu istemcide desteklenmiyor.");
    }

    let platform = "Bilinmiyor";
    // Platform enum'ını daha okunabilir bir metne çevir
    switch (Office.context.diagnostics.platform) {
        case Office.PlatformType.PC:
            platform = "Windows (Masaüstü)";
            break;
        case Office.PlatformType.Mac:
            platform = "Mac (Masaüstü)";
            break;
        case Office.PlatformType.OfficeOnline:
            platform = "Web (Tarayıcı)";
            break;
        case Office.PlatformType.iOS:
            platform = "iOS (Mobil)";
            break;
        case Office.PlatformType.Android:
            platform = "Android (Mobil)";
            break;
        default:
            platform = "Diğer";
            break;
    }
    
    // Gerekli bilgileri bir nesne içinde topla
    const clientInfo = {
        host: Office.HostType[Office.context.diagnostics.host], // "Outlook" döner
        applicationVersion: Office.context.diagnostics.version, // Örn: "16.0.12345.98765"
        platform: platform
    };

    return clientInfo;
}

/**
 * Geliştiriciler için daha faydalı bir kontrol: Belirli bir API setinin
 * (Requirement Set) desteklenip desteklenmediğini kontrol eder.
 */
function checkApiSupport() {
    const apiSupportElement = document.getElementById("api-support-info");
    
    // Check IdentityAPI support
    const isIdentityApiSupported = Office.context.requirements.isSetSupported('IdentityAPI', '1.3');

    if (isIdentityApiSupported) {
        apiSupportElement.innerHTML = `<p style="color: green;">✅ Kimlik Doğrulama API'si (IdentityAPI 1.3) bu platformda <strong>destekleniyor</strong>.</p>`;
    } else {
        apiSupportElement.innerHTML = `<p style="color: orange;">❌ Kimlik Doğrulama API'si (IdentityAPI 1.3) bu platformda <strong>desteklenmiyor</strong>.</p>`;
    }
    
    // Check authentication status
    checkAuthenticationStatus();
}

// Check and display authentication status
async function checkAuthenticationStatus() {
    const authStatusElement = document.getElementById("auth-status-info");
    
    try {
        // Try to get access token without prompting user
        const token = await authManager.getAccessToken({ allowSignInPrompt: false });
        
        if (token) {
            authStatusElement.innerHTML = `
                <p style="color: green;">✅ <strong>Kimlik doğrulama başarılı</strong></p>
                <p style="font-size: 12px; color: #605e5c;">Microsoft Graph API'ye erişim sağlandı</p>
                <button onclick="testGraphConnection()" style="padding: 4px 8px; font-size: 12px; margin-top: 8px;">Bağlantıyı Test Et</button>
            `;
        }
    } catch (error) {
        console.log('Authentication check failed:', error);
        
        let errorMessage = 'Kimlik doğrulama gerekli';
        let errorColor = 'orange';
        let actionButton = '';
        
        if (error.message.includes('not supported')) {
            errorMessage = 'SSO bu platformda desteklenmiyor';
            errorColor = 'red';
        } else if (error.message.includes('not signed in')) {
            errorMessage = 'Office\'e giriş yapılmamış';
            actionButton = '<button onclick="attemptSignIn()" style="padding: 4px 8px; font-size: 12px; margin-top: 8px;">Giriş Yap</button>';
        } else {
            actionButton = '<button onclick="attemptSignIn()" style="padding: 4px 8px; font-size: 12px; margin-top: 8px;">Kimlik Doğrula</button>';
        }
        
        authStatusElement.innerHTML = `
            <p style="color: ${errorColor};">⚠️ <strong>${errorMessage}</strong></p>
            <p style="font-size: 12px; color: #605e5c;">${error.message}</p>
            ${actionButton}
        `;
    }
}

// Test Graph API connection
async function testGraphConnection() {
    const authStatusElement = document.getElementById("auth-status-info");
    
    try {
        authStatusElement.innerHTML = '<p>🔄 Bağlantı test ediliyor...</p>';
        
        const userProfile = await getUserProfile();
        
        authStatusElement.innerHTML = `
            <p style="color: green;">✅ <strong>Graph API bağlantısı başarılı</strong></p>
            <p style="font-size: 12px; color: #605e5c;">Kullanıcı: ${userProfile.displayName}</p>
            <p style="font-size: 12px; color: #605e5c;">E-posta: ${userProfile.emailAddress}</p>
            <button onclick="checkAuthenticationStatus()" style="padding: 4px 8px; font-size: 12px; margin-top: 8px;">Yenile</button>
        `;
    } catch (error) {
        authStatusElement.innerHTML = `
            <p style="color: red;">❌ <strong>Bağlantı testi başarısız</strong></p>
            <p style="font-size: 12px; color: #605e5c;">${error.message}</p>
            <button onclick="checkAuthenticationStatus()" style="padding: 4px 8px; font-size: 12px; margin-top: 8px;">Tekrar Dene</button>
        `;
    }
}

// Attempt to sign in
async function attemptSignIn() {
    const authStatusElement = document.getElementById("auth-status-info");
    
    try {
        authStatusElement.innerHTML = '<p>🔄 Giriş yapılıyor...</p>';
        
        const token = await authManager.getAccessToken({ allowSignInPrompt: true });
        
        if (token) {
            authStatusElement.innerHTML = `
                <p style="color: green;">✅ <strong>Giriş başarılı</strong></p>
                <p style="font-size: 12px; color: #605e5c;">Microsoft Graph API'ye erişim sağlandı</p>
                <button onclick="testGraphConnection()" style="padding: 4px 8px; font-size: 12px; margin-top: 8px;">Bağlantıyı Test Et</button>
            `;
        }
    } catch (error) {
        authStatusElement.innerHTML = `
            <p style="color: red;">❌ <strong>Giriş başarısız</strong></p>
            <p style="font-size: 12px; color: #605e5c;">${error.message}</p>
            <button onclick="checkAuthenticationStatus()" style="padding: 4px 8px; font-size: 12px; margin-top: 8px;">Tekrar Dene</button>
        `;
    }
}

// Fallback initialization for testing in browser
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM loaded, initializing app');
    // Add a small delay to ensure Office.js has time to load
    setTimeout(() => {
        if (!window.officeInitialized) {
            console.log('Office.js not initialized, using fallback');
            initializeApp();
        }
    }, 1000);
});

function initializeApp() {
    console.log('Initializing app');
    window.officeInitialized = true;
    
    const form = document.getElementById('autoReplyForm');
    if (form) {
        form.addEventListener('submit', setAutoReply);
    }
    
    checkCurrentOofStatus();
    setupColleagueSearch();
    setupFormListeners();
    setDefaultDates();
    updateVersionInfo();
}

// Global variables
let colleagues = [];
let selectedColleague = null;

// Message template (both Turkish and English)
const messageTemplate = {
    subject: "Otomatik Yanıt: Yıllık İzin / Automatic Reply: Annual Leave",
    body: `Sayın Yetkili,

E-postanız için teşekkür ederim. {startDate} – {endDate} tarihleri arasında yıllık izinde olacağım ve bu süre içinde e-postalarınıza yanıt veremeyeceğim.

Acil konularınız için {colleagueName} ile {email} {phone} üzerinden iletişime geçebilirsiniz.

Anlayışınız için teşekkür eder, iyi çalışmalar dilerim.

Saygılarımla,
{userName}
{position}
{company}

---

Dear Sir/Madam,

Thank you for your email. I will be out of the office on annual leave from {startDate} to {endDate}, and will not be able to respond to your message during this period.

For urgent matters, please contact {colleagueName} at {email} {phoneEn}.

Thank you for your understanding.

Kind regards,
{userName}
{position}
{company}`
};

// Check current OOF status using Graph API with new authentication
async function checkCurrentOofStatus() {
    try {
        console.log('Checking current OOF status...');
        
        const token = await authManager.getAccessToken({ allowSignInPrompt: false });
        
        const response = await fetch('https://graph.microsoft.com/v1.0/me/mailboxSettings', {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            const oofSettings = data.automaticRepliesSetting;
            
            console.log('Current OOF settings:', oofSettings);
            
            if (oofSettings && (oofSettings.status === 'enabled' || oofSettings.status === 'scheduled')) {
                showOofStatusBanner(oofSettings);
            }
        } else {
            console.error('Failed to get mailbox settings:', response.status, response.statusText);
        }
    } catch (error) {
        console.log('Could not check OOF status:', error);
        // Silently fail - not critical for app functionality
    }
}

// Show OOF status banner
function showOofStatusBanner(oofSettings) {
    const banner = document.getElementById('oofStatusBanner');
    const text = document.getElementById('oofStatusText');
    
    if (banner && text) {
        let statusText = '📧 Otomatik yanıt şu anda aktif';
        
        if (oofSettings.status === 'scheduled' && oofSettings.scheduledEndDateTime) {
            const endDate = new Date(oofSettings.scheduledEndDateTime.dateTime);
            const endDateStr = endDate.toLocaleDateString('tr-TR');
            statusText += ` (${endDateStr} tarihine kadar)`;
        }
        
        text.textContent = statusText;
        //banner.style.display = 'block';
        banner.classList.add('active');
    }
}

// Setup colleague search functionality
async function setupColleagueSearch() {
    const colleagueInput = document.getElementById('colleague');
    const dropdown = document.getElementById('colleagueDropdown');
    
    if (!colleagueInput || !dropdown) return;
    
    colleagueInput.addEventListener('input', async (e) => {
        const query = e.target.value.trim();
        
        if (query.length < 2) {
            dropdown.style.display = 'none';
            return;
        }
        
        try {
            const users = await searchUsers(query);
            displayUserResults(users, dropdown, colleagueInput);
        } catch (error) {
            console.error('Error searching users:', error);
            // Fallback to mock data
            const filtered = mockColleagues.filter(c => 
                c.name.toLowerCase().includes(query.toLowerCase())
            );
            displayUserResults(filtered, dropdown, colleagueInput);
        }
    });
    
    // Hide dropdown when clicking outside
    document.addEventListener('click', (e) => {
        if (!colleagueInput.contains(e.target) && !dropdown.contains(e.target)) {
            dropdown.style.display = 'none';
        }
    });
}

async function getAccessToken() {
    try {
        const token = "eyJ0eXAiOiJKV1QiLCJub25jZSI6ImI2ZHpWMVNWNFFJWkpPN0REcHdoN05IUFZ2b0pvTWZtNnJtOXNaUmhhOGsiLCJhbGciOiJSUzI1NiIsIng1dCI6IkpZaEFjVFBNWl9MWDZEQmxPV1E3SG4wTmVYRSIsImtpZCI6IkpZaEFjVFBNWl9MWDZEQmxPV1E3SG4wTmVYRSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8yNjM3MWI1ZS05NDhlLTQ3OTEtYTA1NS0wMDNjOTAzMDZiOGMvIiwiaWF0IjoxNzU3Njg0MzU3LCJuYmYiOjE3NTc2ODQzNTcsImV4cCI6MTc1NzY4ODI1NywiYWlvIjoiazJSZ1lGZzA2eTduMjdaTFVvRzNJeHJlM0t4UkFnQT0iLCJhcHBfZGlzcGxheW5hbWUiOiJPdXRsb29rIE9PRiIsImFwcGlkIjoiYzJhOGI2NTAtNTBiMi00NDZlLWE4ZTktYmZmYTY2OThiNzdmIiwiYXBwaWRhY3IiOiIxIiwiaWRwIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvMjYzNzFiNWUtOTQ4ZS00NzkxLWEwNTUtMDAzYzkwMzA2YjhjLyIsImlkdHlwIjoiYXBwIiwib2lkIjoiMGNjNDdmYzUtYWQ1Mi00MTFlLTliN2QtYmU3ZThjYmQ3NjcxIiwicmgiOiIxLkFWd0FYaHMzSm82VWtVZWdWUUE4a0RCcmpBTUFBQUFBQUFBQXdBQUFBQUFBQUFCY0FBQmNBQS4iLCJyb2xlcyI6WyJQZW9wbGUuUmVhZC5BbGwiLCJVc2VyLlJlYWQuQWxsIl0sInN1YiI6IjBjYzQ3ZmM1LWFkNTItNDExZS05YjdkLWJlN2U4Y2JkNzY3MSIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJFVSIsInRpZCI6IjI2MzcxYjVlLTk0OGUtNDc5MS1hMDU1LTAwM2M5MDMwNmI4YyIsInV0aSI6Ik9hSVlDdXlVMWtpWVFNQjVGS0VJQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjA5OTdhMWQwLTBkMWQtNGFjYi1iNDA4LWQ1Y2E3MzEyMWU5MCJdLCJ4bXNfZnRkIjoicmREalJEbG1sR1RwRkxqd0thVGtpLW1GS0dwZG9TN3hCeWhwN2oxYnN1c0JaWFZ5YjNCbGQyVnpkQzFrYzIxeiIsInhtc19pZHJlbCI6IjI0IDciLCJ4bXNfcmQiOiIwLjQyTGxZQkppakJFUzRXQVhFZ2pPdmpoenpfTXJudlAydmRvaHRYdTdMVkNVVTBoZzlnZi1mWk12TWpuMTdIcXYtUFhHc1MxQVVRNGhBV1lHQ0RnQXBRRSIsInhtc190Y2R0IjoxNDQ5MjI3NzY4fQ.RUwb1Uk8qzeEf6RZo0Sy6MzSZnUNDKKKRi8vcL_u94Xmqsu0ow3fyxgxejIZF6jhufOJTdHa_GMdVpfXSsqncwT5eBG9vY-tC6Y2Sgs6jDsLEj7_7DnN3GJStlLAGYZebBk_B-DFTRm0RwmFrDiP8U_mtzXXDDdKFxyswo0UanW8MgVOTeN5SIzCTqoME7bjkAtkLSyjs33g2CScl-Zye8KXB9X_4WbfEcQ5WNkDg2Db0ruYyrMW5_wGA6-TtUpf0WBu-fc-ZwS4eBY1NagcZlR3T9MnGJyVMY6fyDacIAP9s5CHxnojBDJwFlYZ-KKKtNE_wMnkR3cD9buf5Gv9lw";
        return token;
    } catch (error) {
        console.error('Error getting access token:', error);
        return null;
    }
}

// Search users via Graph API with new authentication
async function searchUsers(query) {
    try {
        console.log('Searching users with query:', query);
        
        const token = await getAccessToken();
        
        const encodedQuery = encodeURIComponent(query);
        const response = await fetch(`https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${encodedQuery}') or startswith(givenName,'${encodedQuery}') or startswith(surname,'${encodedQuery}')&$select=id,displayName,mail,userPrincipalName,jobTitle,businessPhones&$top=10`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            console.log('User search results:', data.value);
            
            return data.value.map(user => ({
                id: user.id,
                name: user.displayName,
                email: user.mail || user.userPrincipalName,
                jobTitle: user.jobTitle,
                phone: user.mobilePhone
            }));
        } else {
            console.error('User search failed:', response.status, response.statusText);
            throw new Error(`User search failed: ${response.status}`);
        }
    } catch (error) {
        console.error('Graph API user search failed:', error);
        throw error;
    }
    
    return [];
}

// Display user search results
function displayUserResults(users, dropdown, input) {
    dropdown.innerHTML = '';
    
    if (users.length === 0) {
        dropdown.innerHTML = '<div class="colleague-item">Kullanıcı bulunamadı</div>';
        dropdown.style.display = 'block';
        return;
    }
    
    users.forEach(user => {
        const item = document.createElement('div');
        item.className = 'colleague-item';
        item.innerHTML = `
            <div><strong>${user.name}</strong></div>
            <div style="font-size: 12px; color: #605e5c;">${user.jobTitle} - ${user.email}</div>
        `;
        
        item.addEventListener('click', () => {
            selectedColleague = user;
            input.value = user.name;
            dropdown.style.display = 'none';
            updatePreview().catch(error => {
                console.error('Error updating preview:', error);
            });
        });
        
        dropdown.appendChild(item);
    });
    
    dropdown.style.display = 'block';
}

function setupFormListeners() {
    const inputs = ['colleague', 'startDate', 'startTime', 'endDate', 'endTime'];
    inputs.forEach(id => {
        document.getElementById(id).addEventListener('change', () => {
            updatePreview().catch(error => {
                console.error('Error updating preview:', error);
            });
        });
    });
}

function setDefaultDates() {
    const today = new Date();
    const nextWeek = new Date(today);
    nextWeek.setDate(nextWeek.getDate() + 7);
    
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');
    
    // Set minimum dates
    startDateInput.min = formatDate(today);
    endDateInput.min = formatDate(today);
    
    // Set default values
    startDateInput.value = formatDate(today);
    endDateInput.value = formatDate(nextWeek);
    
    // Add event listener to update end date minimum when start date changes
    startDateInput.addEventListener('change', function() {
        const startDate = new Date(this.value);
        endDateInput.min = formatDate(startDate);
        
        // If end date is before start date, update it
        if (endDateInput.value && new Date(endDateInput.value) < startDate) {
            endDateInput.value = formatDate(startDate);
        }
    });
}

function formatDate(date) {
    return date.toISOString().split('T')[0];
}

function formatDisplayDate(dateStr, timeStr) {
    const date = new Date(dateStr + 'T' + timeStr);
    
    return date.toLocaleDateString('tr-TR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    }) + ' ' + timeStr;
}

async function updatePreview() {
    const colleagueInput = document.getElementById('colleague').value;
    const startDate = document.getElementById('startDate').value;
    const startTime = document.getElementById('startTime').value;
    const endDate = document.getElementById('endDate').value;
    const endTime = document.getElementById('endTime').value;
    
    const previewDiv = document.getElementById('messagePreview');
    
    if (!selectedColleague || !startDate || !startTime || !endDate || !endTime) {
        previewDiv.textContent = 'Lütfen tüm alanları doldurun...';
        return;
    }
    
    const startDateTime = formatDisplayDate(startDate, startTime);
    const endDateTime = formatDisplayDate(endDate, endTime);
    
    // Get current user info using the proper getUserProfile function
    let currentUser;
    try {
        currentUser = await getUserProfile();
    } catch (error) {
        console.error('Error getting user profile:', error);
        // Fallback to Office context if getUserProfile fails
        currentUser = {
            name: Office.context.mailbox.userProfile?.displayName || 'Kullanıcı Adı',
            position: Office.context.mailbox.userProfile?.jobTitle || '',
            company: Office.context.mailbox.userProfile?.companyName || 'Öztiryakiler'
        };
    }
    
    let messageBody = messageTemplate.body
        .replaceAll('{startDate}', startDateTime)
        .replaceAll('{endDate}', endDateTime)
        .replaceAll('{colleagueName}', selectedColleague.name)
        .replaceAll('{email}', selectedColleague.email)
        .replaceAll('{phone}', " veya " + selectedColleague.phone || '')
        .replaceAll('{phoneEn}', " or " + selectedColleague.phone || '')
        .replaceAll('{userName}', currentUser.displayName || currentUser.name)
        .replaceAll('{position}', currentUser.jobTitle || currentUser.position)
        .replaceAll('{company}', currentUser.companyName || currentUser.company);
    
    previewDiv.textContent = `Konu: ${messageTemplate.subject}

${messageBody}`;
}

async function setAutoReply(event) {
    event.preventDefault();
    
    const colleagueInput = document.getElementById('colleague').value;
    const startDate = document.getElementById('startDate').value;
    const startTime = document.getElementById('startTime').value;
    const endDate = document.getElementById('endDate').value;
    const endTime = document.getElementById('endTime').value;
    
    if (!selectedColleague || !startDate || !startTime || !endDate || !endTime) {
        showStatus('error', 'Lütfen tüm alanları doldurun!');
        return;
    }
    
    const startDateTime = new Date(startDate + 'T' + startTime);
    const endDateTime = new Date(endDate + 'T' + endTime);
    
    if (startDateTime >= endDateTime) {
        showStatus('error', 'Bitiş tarihi başlangıç tarihinden sonra olmalıdır!');
        return;
    }
    
    const button = document.getElementById('btnSetAutoReply');
    button.disabled = true;
    button.textContent = 'Ayarlanıyor...';
    
    try {
        // Get current user information
        const userProfile = await getUserProfile();
        
        // Prepare the auto-reply message
        const startDateTimeFormatted = formatDisplayDate(startDate, startTime);
        const endDateTimeFormatted = formatDisplayDate(endDate, endTime);
        
        let messageBody = messageTemplate.body
            .replaceAll('{startDate}', startDateTimeFormatted)
            .replaceAll('{endDate}', endDateTimeFormatted)
            .replaceAll('{colleagueName}', selectedColleague.name)
            .replaceAll('{email}', selectedColleague.email)
            .replaceAll('{phone}', selectedColleague.phone)
            .replaceAll('{userName}', userProfile.displayName || 'Kullanıcı')
            .replaceAll('{position}', userProfile.jobTitle || 'Pozisyon')
            .replaceAll('{company}', userProfile.companyName || 'Öztiryakiler');
        
        // Set the automatic reply using Graph API
        console.log('setOutlookAutoReply before');
        await setOutlookAutoReply(messageBody, startDateTime, endDateTime);
        
        showStatus('success', 'Otomatik yanıt mesajı hazırlandı! Manuel ayarlama talimatları gösterilecek.');
        
        // Log the auto-reply details for debugging
        console.log('Auto-reply set:', {
            subject: messageTemplate.subject,
            body: messageBody,
            startDate: startDateTime,
            endDate: endDateTime,
            colleague: colleague
        });
        
    } catch (error) {
        console.error('Error setting auto-reply:', error);
        showStatus('error', 'Otomatik yanıt ayarlanırken hata oluştu: ' + error.message);
    } finally {
        button.disabled = false;
        button.textContent = 'Otomatik Yanıtı Ayarla';
    }
}

// Get user profile information using Microsoft Graph API with new authentication
async function getUserProfile() {
    try {
        console.log('Getting user profile...');
        
        const token = await authManager.getAccessToken();

        const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,companyName", {
            headers: {
                "Authorization": "Bearer " + token,
                "Content-Type": "application/json"
            }
        });

        if (response.ok) {
            const userData = await response.json();
            
            console.log('Graph API user data:', userData);
            return {
                displayName: userData.displayName || 'Kullanıcı',
                emailAddress: userData.mail || userData.userPrincipalName || 'user@oztiryakiler.com.tr',
                jobTitle: userData.jobTitle || '',
                companyName: userData.companyName || 'Öztiryakiler'
            };
        } else {
            console.error("Graph API user profile request failed:", response.status, response.statusText);
            throw new Error(`Graph API failed: ${response.status}`);
        }

    } catch (exception) {
        console.error("Graph API error, falling back to Office context:", exception);
        
        // Fallback to Office context
        if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) {
            const userProfile = Office.context.mailbox.userProfile;
            console.log('Using Office context user profile:', userProfile.displayName);
            return {
                displayName: userProfile.displayName || 'Kullanıcı',
                emailAddress: userProfile.emailAddress || 'user@oztiryakiler.com.tr',
                jobTitle: userProfile.jobTitle || '',
                companyName: 'Öztiryakiler'
            };
        } else {
            // Final fallback for testing or when Office context is not available
            console.log('Using fallback user profile');
            return {
                displayName: 'Test Kullanıcısı',
                emailAddress: 'test@oztiryakiler.com.tr',
                jobTitle: '',
                companyName: 'Öztiryakiler'
            };
        }
    }
}

// Set Outlook automatic reply - Try Graph API first, then EWS as fallback
async function setOutlookAutoReply(messageBody, startDateTime, endDateTime) {
    try {
        const result = await setAutomaticReply(startDateTime, endDateTime, messageBody, messageBody);
        console.log(`Auto-reply set successfully using: ${result}`);
        
        if (result === "graph") {
            showStatus('success', 'Otomatik yanıt Graph API ile başarıyla ayarlandı!');
        } else if (result === "ews") {
            showStatus('success', 'Otomatik yanıt EWS ile başarıyla ayarlandı!');
        }
        
    } catch (error) {
        console.error('Both Graph API and EWS failed:', error);
        showStatus('warning', 'Otomatik ayarlama başarısız oldu. Manuel talimatlar gösteriliyor.');
        showInstructions(messageBody, startDateTime, endDateTime);
    }
}

// Enhanced automatic reply setting with new authentication
async function setAutomaticReply(startLocal, endLocal, internalMsg, externalMsg) {
  try {
    console.log('Setting automatic reply via Graph API...');
    
    const token = await authManager.getAccessToken();
    await setOOFViaGraph(token, startLocal, endLocal, internalMsg, externalMsg);
    
    // Verify the setting was applied
    await new Promise(resolve => setTimeout(resolve, 2000)); // Wait 2 seconds
    const verification = await verifyOOFViaGraph(token);
    
    if (verification && (verification.status === 'scheduled' || verification.status === 'enabled')) {
        console.log('Graph API OOF setting verified successfully');
        return "graph";
    } else {
        console.log('Graph API verification failed, trying EWS fallback');
        throw new Error('Graph API verification failed');
    }
    
  } catch (graphError) {
    console.log('Graph API failed, trying EWS:', graphError);
    
    // Try EWS with SET → GET verification
    try {
      await setOOFViaEws(startLocal, endLocal, internalMsg, externalMsg, "All");
      
      // Verify with GET request
      const verification = await getOOFViaEws();
      
      if (verification.state === "Scheduled") {
        console.log('EWS verification successful:', verification);
        return "ews";
      } else {
        throw new Error(`EWS verification failed: state=${verification.state}`);
      }
    } catch (ewsError) {
      console.error('EWS also failed:', ewsError);
      throw ewsError;
    }
  }
}

// Verify OOF setting via Graph API
async function verifyOOFViaGraph(token) {
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/mailboxSettings/automaticRepliesSetting', {
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json'
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            console.log('OOF verification result:', data);
            return data;
        }
    } catch (error) {
        console.error('OOF verification failed:', error);
    }
    return null;
}

// 2) GRAPH (hazır olduğunda)
async function setOOFViaGraph(token, startLocal, endLocal, internalMsg, externalMsg) {
  const toISO = d => new Date(d).toISOString().slice(0,19); // yyyy-MM-ddTHH:mm:ss
  const body = {
    automaticRepliesSetting: {
      status: "scheduled",
      scheduledStartDateTime: { dateTime: toISO(startLocal), timeZone: "Turkey Standard Time" },
      scheduledEndDateTime:   { dateTime: toISO(endLocal),   timeZone: "Turkey Standard Time" },
      internalReplyMessage: internalMsg,
      externalReplyMessage: externalMsg
    }
  };
  const res = await fetch("https://graph.microsoft.com/v1.0/me/mailboxSettings", {
    method: "PATCH",
    headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(await res.text());
}

// Helper functions for EWS
function toLocalNaive(dt) {
  // dt: Date veya "2025-09-10T07:30" gibi yerel saat
  const d = new Date(dt);
  // Yereldir: Z/yükleme yok, saniye ekleyelim
  const pad = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:00`;
}

function xmlEscape(s) {
  return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

// 3) EWS – ANINDA çalışır (Improved version with diagnostics)
function toLocalNaive(dt) {
    const d = new Date(dt);
    const p = n => String(n).padStart(2,'0');
    return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:00`;
  }
  
  // ]]>, CDATA’yı bozmasın:
  function cdataWrap(s) {
    return '<![CDATA[' + String(s).replace(/]]>/g, ']]]]><![CDATA[>') + ']]>';
  }
  

  /**
 * Kullanıcının OOF ayarlarını güncelleyen asenkron fonksiyon
 * @param {boolean} isOofEnabled - OOF ayarının açık mı kapalı mı olacağını belirtir.
 * @param {Date} startTime - OOF başlangıç tarihi. (İsteğe bağlı)
 * @param {Date} endTime - OOF bitiş tarihi. (İsteğe bağlı)
 * @param {string} internalMessage - Organizasyon içi otomatik yanıt mesajı.
 * @param {string} externalMessage - Organizasyon dışı otomatik yanıt mesajı.
 */
async function updateOofSettings(isOofEnabled, startTime, endTime, internalMessage, externalMessage) {
    // EWS URL'sinin mevcut olduğunu kontrol edin
    if (!Office.context.mailbox.ewsUrl) {
        console.error("Exchange Web Services URL'si mevcut değil.");
        return;
    }

    const OofState = isOofEnabled ? 'Scheduled' : 'Disabled';

    const soapRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2016" />
  </soap:Header>
  <soap:Body>
    <SetUserOofSettings xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <Mailbox>
        <t:Address>
          <t:Name>${Office.context.mailbox.userProfile.displayName}</t:Name>
          <t:EmailAddress>${Office.context.mailbox.userProfile.emailAddress}</t:EmailAddress>
        </t:Address>
      </Mailbox>
      <UserOofSettings>
        <t:OofState>${OofState}</t:OofState>
        <t:ExternalAudience>All</t:ExternalAudience>
        ${isOofEnabled && startTime && endTime ? 
            `<t:Duration>
                <t:StartTime>${new Date(startTime).toISOString()}</t:StartTime>
                <t:EndTime>${new Date(endTime).toISOString()}</t:EndTime>
            </t:Duration>`
            : ''}
        <t:InternalReply>
          <t:Message>${internalMessage}</t:Message>
        </t:InternalReply>
        <t:ExternalReply>
          <t:Message>${externalMessage}</t:Message>
        </t:ExternalReply>
      </UserOofSettings>
    </SetUserOofSettings>
  </soap:Body>
</soap:Envelope>`;

    try {
        const response = await fetch(Office.context.mailbox.ewsUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'text/xml; charset=utf-8',
                'SOAPAction': '"http://schemas.microsoft.com/exchange/services/2006/messages/SetUserOofSettings"',
                // Office.js'de kimlik doğrulama, Office tarafından otomatik olarak yönetilir,
                // bu yüzden manuel bir Authorization başlığına genellikle gerek yoktur.
            },
            body: soapRequest
        });

        if (!response.ok) {
            throw new Error(`HTTP hatası! Durum: ${response.status}`);
        }

        const responseText = await response.text();
        console.log("OOF ayarı başarıyla güncellendi. Yanıt:", responseText);

    } catch (error) {
        console.error("OOF ayarı güncellenirken bir hata oluştu:", error);
    }
}


  function setOOFViaEws(startLocal, endLocal, internalMsg, externalMsg, audience = "All") {
    //console.log("token",Office.context.mailbox.getAccessToken());
    
    const email = Office.context.mailbox.userProfile.emailAddress;
    const start = toLocalNaive(startLocal);
    const end   = toLocalNaive(endLocal);
  

    updateOofSettings(true, start, end, internalMsg, internalMsg);

    // const soap = `
    // <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
    //                xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
    //                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    //   <soap:Header>
    //     <t:RequestServerVersion Version="Exchange2013"/>
    //   </soap:Header>
    //   <soap:Body>
    //     <m:SetUserOofSettingsRequest>
    //       <t:Mailbox><t:Address>${email}</t:Address></t:Mailbox>
    //       <t:UserOofSettings>
    //         <t:OofState>Scheduled</t:OofState>
    //         <t:ExternalAudience>${audience}</t:ExternalAudience>
    //         <t:Duration>
    //           <t:StartTime>${start}</t:StartTime>
    //           <t:EndTime>${end}</t:EndTime>
    //         </t:Duration>
    //         <t:InternalReply><t:Message>${cdataWrap(internalMsg)}</t:Message></t:InternalReply>
    //         <t:ExternalReply><t:Message>${cdataWrap(externalMsg)}</t:Message></t:ExternalReply>
    //       </t:UserOofSettings>
    //     </m:SetUserOofSettingsRequest>
    //   </soap:Body>
    // </soap:Envelope>`;
  
    // const M = "http://schemas.microsoft.com/exchange/services/2006/messages";
  
    // return new Promise((resolve, reject) => {
    //   Office.context.mailbox.makeEwsRequestAsync(soap, res => {
    //     if (res.status !== Office.AsyncResultStatus.Succeeded) {
    //       return reject(res.error);
    //     }
    //     console.log("setOOFViaEws result",res)
    //     const xml = new DOMParser().parseFromString(res.value, "text/xml");
    //     const code = xml.getElementsByTagNameNS(M, "ResponseCode")[0]?.textContent;
    //     const text = xml.getElementsByTagNameNS(M, "MessageText")[0]?.textContent || "";
    //     if (code === "NoError") return resolve();
    //     reject(new Error(`EWS ResponseCode: ${code || "Unknown"} ${text}`.trim()));
    //   });
    // });
  }
  
  // (İsteğe bağlı) anında doğrulama:
  function getOOFViaEws() {
    const email = Office.context.mailbox.userProfile.emailAddress;
    const soap = `
    <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                   xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <soap:Body>
        <m:GetUserOofSettingsRequest>
          <t:Mailbox><t:Address>${email}</t:Address></t:Mailbox>
        </m:GetUserOofSettingsRequest>
      </soap:Body>
    </soap:Envelope>`;
    const T = "http://schemas.microsoft.com/exchange/services/2006/types";
    return new Promise((resolve, reject) => {
      Office.context.mailbox.makeEwsRequestAsync(soap, res => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(res.error);
        const xml = new DOMParser().parseFromString(res.value, "text/xml");
        resolve({
          state: xml.getElementsByTagNameNS(T, "OofState")[0]?.textContent,
          start: xml.getElementsByTagNameNS(T, "StartTime")[0]?.textContent,
          end:   xml.getElementsByTagNameNS(T, "EndTime")[0]?.textContent,
          raw:   res.value
        });
      });
    });
  }
  

// // GET function to verify OOF settings
// function getOOFViaEws() {
//   const email = Office.context.mailbox.userProfile.emailAddress;
//   const soap = `
//   <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
//     <soap:Body>
//       <GetUserOofSettingsRequest xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
//         <Mailbox xmlns="http://schemas.microsoft.com/exchange/services/2006/types">${email}</Mailbox>
//       </GetUserOofSettingsRequest>
//     </soap:Body>
//   </soap:Envelope>`;
  
//   return new Promise((resolve, reject) => {
//     Office.context.mailbox.makeEwsRequestAsync(soap, res => {
//       if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(res.error);
//       const xml = new window.DOMParser().parseFromString(res.value, "text/xml");
//       const state = xml.getElementsByTagName("OofState")[0]?.textContent;
//       const start = xml.getElementsByTagName("StartTime")[0]?.textContent;
//       const end   = xml.getElementsByTagName("EndTime")[0]?.textContent;
//       const ext   = xml.getElementsByTagName("ExternalAudience")[0]?.textContent;
      
//       console.log('EWS GET Results:', { state, start, end, ext });
//       console.log('EWS GET Response XML:', res.value);
      
//       resolve({ state, start, end, ext, raw: res.value });
//     });
//   });
// }

// Helper function to escape XML characters
function escapeXml(text) {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// Set auto-reply via Microsoft Graph API (kept for reference)
async function setAutoReplyViaGraphAPI(accessToken, messageBody, startDateTime, endDateTime) {
    const graphEndpoint = 'https://graph.microsoft.com/v1.0/me/mailboxSettings';
    
    const autoReplySettings = {
        automaticRepliesSetting: {
            status: 'scheduled',
            externalAudience: 'all',
            scheduledStartDateTime: {
                dateTime: startDateTime.toISOString(),
                timeZone: 'Turkey Standard Time'
            },
            scheduledEndDateTime: {
                dateTime: endDateTime.toISOString(),
                timeZone: 'Turkey Standard Time'
            },
            internalReplyMessage: messageBody,
            externalReplyMessage: messageBody
        }
    };
    
    const response = await fetch(graphEndpoint, {
        method: 'PATCH',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(autoReplySettings)
    });
    
    console.log('setAutoReplyViaGraphAPI accessToken:' , accessToken);
    console.log('setAutoReplyViaGraphAPI json:' , autoReplySettings);
    console.log('setAutoReplyViaGraphAPI response:' , response);

    if (!response.ok) {
        throw new Error(`Graph API error: ${response.status}`);
    }
    
    return response;
}

// Show manual instructions to user
function showInstructions(messageBody, startDateTime, endDateTime) {
    const modal = document.getElementById('instructionsModal');
    const content = document.getElementById('instructionsContent');
    
    const startDateStr = startDateTime.toLocaleDateString('tr-TR') + ' ' + startDateTime.toLocaleTimeString('tr-TR', {hour: '2-digit', minute: '2-digit'});
    const endDateStr = endDateTime.toLocaleDateString('tr-TR') + ' ' + endDateTime.toLocaleTimeString('tr-TR', {hour: '2-digit', minute: '2-digit'});
    
    content.innerHTML = `
        <div class="instruction-step">
            <strong>1. Outlook Ayarlarını Açın</strong><br>
            Dosya → Otomatik Yanıtlar (Ofis Dışında) menüsüne gidin.
        </div>
        
        <div class="instruction-step">
            <strong>2. Otomatik Yanıtları Etkinleştirin</strong><br>
            "Otomatik yanıtları gönder" seçeneğini işaretleyin.
        </div>
        
        <div class="instruction-step">
            <strong>3. Zaman Aralığını Ayarlayın</strong><br>
            "Yalnızca şu zaman aralığında gönder" seçeneğini işaretleyin:<br>
            <strong>Başlangıç:</strong> ${startDateStr}<br>
            <strong>Bitiş:</strong> ${endDateStr}
        </div>
        
        <div class="instruction-step">
            <strong>4. Mesaj İçeriğini Kopyalayın</strong><br>
            Aşağıdaki mesajı kopyalayıp "Kuruluşum içinde" ve "Kuruluşum dışında" alanlarına yapıştırın:
            <button class="copy-button" onclick="copyMessage()">📋 Kopyala</button>
            <div id="messageForCopy" style="display: none;">${messageBody}</div>
        </div>
        
        <div class="instruction-step">
            <strong>5. Kaydedin</strong><br>
            "Tamam" butonuna tıklayarak ayarları kaydedin.
        </div>
    `;
    
    modal.style.display = 'block';
}

// Copy message to clipboard
function copyMessage() {
    const messageDiv = document.getElementById('messageForCopy');
    const textArea = document.createElement('textarea');
    textArea.value = messageDiv.textContent;
    document.body.appendChild(textArea);
    textArea.select();
    document.execCommand('copy');
    document.body.removeChild(textArea);
    
    // Show feedback
    const copyButton = event.target;
    const originalText = copyButton.textContent;
    copyButton.textContent = '✅ Kopyalandı!';
    setTimeout(() => {
        copyButton.textContent = originalText;
    }, 2000);
}

// Copy preview message to clipboard with 30-second feedback
function copyPreviewMessage() {
    const messagePreview = document.getElementById('messagePreview');
    const copyButton = document.getElementById('copyPreviewButton');
    
    // Create a temporary textarea to copy the text
    const textArea = document.createElement('textarea');
    textArea.value = messagePreview.textContent;
    document.body.appendChild(textArea);
    textArea.select();
    document.execCommand('copy');
    document.body.removeChild(textArea);
    
    // Show feedback with checkmark
    const originalText = copyButton.innerHTML;
    copyButton.innerHTML = '✅ Kopyalandı!';
    
    // Revert back to original text after 30 seconds
    setTimeout(() => {
        copyButton.innerHTML = originalText;
    }, 30000);
}

// Close instructions modal
function closeInstructions() {
    document.getElementById('instructionsModal').style.display = 'none';
}

function showStatus(type, message) {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.className = `status-message status-${type}`;
    statusDiv.textContent = message;
    statusDiv.style.display = 'block';
    
    setTimeout(() => {
        statusDiv.style.display = 'none';
    }, 8000); // Show longer for success messages
}

// Update version information
function updateVersionInfo() {
    const versionInfoElement = document.getElementById('versionInfo');
    const lastUpdateElement = document.getElementById('lastUpdate');
    
    if (versionInfoElement) {
        const now = new Date();
        const dateStr = now.toLocaleDateString('tr-TR', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        });
        const timeStr = now.toLocaleTimeString('tr-TR', {
            hour: '2-digit',
            minute: '2-digit'
        });
        
        // Update version info with JavaScript variable
        versionInfoElement.innerHTML = `Sürüm: ${version} | Son Güncelleme: <span id="lastUpdate">${dateStr} ${timeStr}</span>`;
    }
    
    // if (lastUpdateElement) {
    //     const now = new Date();
    //     const dateStr = now.toLocaleDateString('tr-TR', {
    //         day: '2-digit',
    //         month: '2-digit',
    //         year: 'numeric'
    //     });
    //     const timeStr = now.toLocaleTimeString('tr-TR', {
    //         hour: '2-digit',
    //         minute: '2-digit'
    //     });
    //     lastUpdateElement.textContent = `${dateStr} ${timeStr}`;
    // }
}

