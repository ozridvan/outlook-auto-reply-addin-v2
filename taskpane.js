let version = "1.0.1";
// Office.js initialization
console.log('version: '+ version);

Office.onReady((info) => {
    console.log('Office.onReady called', info);
    if (info.host === Office.HostType.Outlook) {
        initializeApp();
    }
});

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
    
    // Örneğin, SSO için gerekli olan IdentityAPI 1.3'ü kontrol edelim
    const isIdentityApiSupported = Office.context.requirements.isSetSupported('IdentityAPI', '1.3');

    if (isIdentityApiSupported) {
        apiSupportElement.innerHTML = `<p style="color: green;">✅ Kimlik Doğrulama API'si (IdentityAPI 1.3) bu platformda <strong>destekleniyor</strong>.</p>`;
    } else {
        apiSupportElement.innerHTML = `<p style="color: orange;">❌ Kimlik Doğrulama API'si (IdentityAPI 1.3) bu platformda <strong>desteklenmiyor</strong>.</p>`;
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

Acil konularınız için {colleagueName} ile {email} veya {phone} üzerinden iletişime geçebilirsiniz.

Anlayışınız için teşekkür eder, iyi çalışmalar dilerim.

Saygılarımla,
{userName}
{position}
{company}

---

Dear Sir/Madam,

Thank you for your email. I will be out of the office on annual leave from {startDate} to {endDate}, and will not be able to respond to your message during this period.

For urgent matters, please contact {colleagueName} at {email} or {phone}.

Thank you for your understanding.

Kind regards,
{userName}
{position}
{company}`
};

// Mock D365 data - In production, this would come from D365 API
const mockColleagues = [
    {
        id: 1,
        name: "Ahmet Yılmaz",
        email: "ahmet.yilmaz@ozturyakiler.com.tr",
        phone: "+90 212 555 0101",
        department: "İnsan Kaynakları"
    },
    {
        id: 2,
        name: "Fatma Demir",
        email: "fatma.demir@ozturyakiler.com.tr",
        phone: "+90 212 555 0102",
        department: "Muhasebe"
    },
    {
        id: 3,
        name: "Mehmet Kaya",
        email: "mehmet.kaya@ozturyakiler.com.tr",
        phone: "+90 212 555 0103",
        department: "Satış"
    },
    {
        id: 4,
        name: "Ayşe Özkan",
        email: "ayse.ozkan@ozturyakiler.com.tr",
        phone: "+90 212 555 0104",
        department: "Pazarlama"
    },
    {
        id: 5,
        name: "Can Şahin",
        email: "can.sahin@ozturyakiler.com.tr",
        phone: "+90 212 555 0105",
        department: "IT"
    },
    {
        id: 6,
        name: "Zeynep Arslan",
        email: "zeynep.arslan@ozturyakiler.com.tr",
        phone: "+90 212 555 0106",
        department: "Hukuk"
    },
    {
        id: 7,
        name: "Murat Çelik",
        email: "murat.celik@ozturyakiler.com.tr",
        phone: "+90 212 555 0107",
        department: "Finans"
    },
    {
        id: 8,
        name: "Elif Koç",
        email: "elif.koc@ozturyakiler.com.tr",
        phone: "+90 212 555 0108",
        department: "Operasyon"
    }
];

// Check current OOF status using Graph API
async function checkCurrentOofStatus() {
    try {
        const token = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: false, 
            allowConsentPrompt: false, 
            forMSGraphAccess: true
        });
        
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
            
            if (oofSettings && (oofSettings.status === 'enabled' || oofSettings.status === 'scheduled')) {
                showOofStatusBanner(oofSettings);
            }
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
        banner.style.display = 'block';
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

// Search users via Graph API
async function searchUsers(query) {
    try {
        const token = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true, 
            allowConsentPrompt: true, 
            forMSGraphAccess: true
        });
        
        const response = await fetch(`https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${query}') or startswith(givenName,'${query}') or startswith(surname,'${query}')&$select=id,displayName,mail,jobTitle,department&$top=10`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            return data.value.map(user => ({
                id: user.id,
                name: user.displayName,
                email: user.mail || user.userPrincipalName,
                department: user.department || 'Bilinmiyor',
                phone: '+90 212 555 0100' // Default phone for now
            }));
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
            <div style="font-size: 12px; color: #605e5c;">${user.department} - ${user.email}</div>
        `;
        
        item.addEventListener('click', () => {
            selectedColleague = user;
            input.value = user.name;
            dropdown.style.display = 'none';
            updatePreview();
        });
        
        dropdown.appendChild(item);
    });
    
    dropdown.style.display = 'block';
}

function setupFormListeners() {
    const inputs = ['colleague', 'startDate', 'startTime', 'endDate', 'endTime'];
    inputs.forEach(id => {
        document.getElementById(id).addEventListener('change', updatePreview);
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

function updatePreview() {
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
    
    // Get current user info (in production, this would come from Office.js)
    const currentUser = {
        name: "Kullanıcı Adı", // This would be retrieved from Office context
        position: "Pozisyon",
        company: "Öztiryakiler"
    };
    
    let messageBody = messageTemplate.body
        .replaceAll('{startDate}', startDateTime)
        .replaceAll('{endDate}', endDateTime)
        .replaceAll('{colleagueName}', selectedColleague.name)
        .replaceAll('{email}', selectedColleague.email)
        .replaceAll('{phone}', selectedColleague.phone)
        .replaceAll('{userName}', currentUser.name)
        .replaceAll('{position}', currentUser.position)
        .replaceAll('{company}', currentUser.company);
    
    previewDiv.textContent = `Konu: ${messageTemplate.subject}\n\n${messageBody}`;
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
            .replaceAll('{company}', 'Öztiryakiler');
        
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

// Get user profile information using Microsoft Graph API
async function getUserProfile() {
    try {
        // 1. Adım: Office'ten SSO token'ını al. 
        // Bu token, Graph API'ye erişim için yeterli değildir, "bootstrap token" olarak geçer.
        // allowSignInPrompt: true -> Gerekirse kullanıcıya oturum açma/onay ekranı gösterir.
        const bootstrapToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });

        // 2. Adım: Alınan token ile Graph API'nin "/me" endpoint'ine istek gönder.
        // Bu endpoint, oturum açmış kullanıcının bilgilerini döndürür.
        const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,jobTitle", {
            headers: {
                "Authorization": "Bearer " + bootstrapToken
            }
        });

        if (response.ok) {
            const userData = await response.json();
            
            // 3. Adım: Gelen veriyi döndür.
            console.log('Graph API user data:', userData);
            return {
                displayName: userData.displayName || 'Kullanıcı',
                emailAddress: userData.mail || 'user@oztiryakiler.com.tr',
                jobTitle: userData.jobTitle || 'Pozisyon'
            };
        } else {
            // Hata durumunu yönet - Office context'e geri dön
            console.error("Graph API isteği başarısız oldu: " + response.status);
            throw new Error('Graph API failed');
        }

    } catch (exception) {
        // Token alma sırasında bir hata oluşursa Office context'i kullan
        console.error("Graph API hatası, Office context'e geçiliyor: " + JSON.stringify(exception));
        
        // Fallback to Office context
        if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) {
            const userProfile = Office.context.mailbox.userProfile;
            console.log('Office context user profile:', userProfile.displayName);
            return {
                displayName: userProfile.displayName || 'Kullanıcı',
                emailAddress: userProfile.emailAddress || 'user@oztiryakiler.com.tr',
                jobTitle: userProfile.jobTitle || 'Pozisyon'
            };
        } else {
            // Final fallback for testing or when Office context is not available
            return {
                displayName: 'Test Kullanıcısı',
                emailAddress: 'test@oztiryakiler.com.tr',
                jobTitle: 'Test Pozisyonu'
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

// 1) GRAPH ile dene, olmazsa EWS'ye düş (with SET → GET verification)
async function setAutomaticReply(startLocal, endLocal, internalMsg, externalMsg) {
  try {
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true
    });
    await setOOFViaGraph(token, startLocal, endLocal, internalMsg, externalMsg);
    return "graph";
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

