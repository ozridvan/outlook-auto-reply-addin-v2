let version = "1.0.0";
// Office.js initialization
console.log('version: '+ version);

Office.onReady((info) => {
    console.log('Office.onReady called', info);
    if (info.host === Office.HostType.Outlook) {
        initializeApp();
    }
});

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
    
    loadColleagues();
    setupFormListeners();
    setDefaultDates();
    updateVersionInfo();
}

// Global variables
let colleagues = [];

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

function loadColleagues() {
    console.log('Loading colleagues...');
    // In production, this would be an API call to D365
    colleagues = mockColleagues;
    console.log('Colleagues loaded:', colleagues.length);
    
    const colleagueSelect = document.getElementById('colleague');
    if (!colleagueSelect) {
        console.error('Colleague select element not found!');
        return;
    }
    
    colleagueSelect.innerHTML = '<option value="">Seçiniz...</option>';
    
    colleagues.forEach(colleague => {
        const option = document.createElement('option');
        option.value = colleague.id;
        option.textContent = `${colleague.name} (${colleague.department})`;
        colleagueSelect.appendChild(option);
        // console.log('Added colleague:', colleague.name);
    });
    
    console.log('Colleagues loaded successfully, total options:', colleagueSelect.options.length);
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
    const colleagueId = document.getElementById('colleague').value;
    const startDate = document.getElementById('startDate').value;
    const startTime = document.getElementById('startTime').value;
    const endDate = document.getElementById('endDate').value;
    const endTime = document.getElementById('endTime').value;
    
    const previewDiv = document.getElementById('messagePreview');
    
    if (!colleagueId || !startDate || !startTime || !endDate || !endTime) {
        previewDiv.textContent = 'Lütfen tüm alanları doldurun...';
        return;
    }
    
    const colleague = colleagues.find(c => c.id == colleagueId);
    
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
        .replaceAll('{colleagueName}', colleague.name)
        .replaceAll('{email}', colleague.email)
        .replaceAll('{phone}', colleague.phone)
        .replaceAll('{userName}', currentUser.name)
        .replaceAll('{position}', currentUser.position)
        .replaceAll('{company}', currentUser.company);
    
    previewDiv.textContent = `Konu: ${messageTemplate.subject}\n\n${messageBody}`;
}

async function setAutoReply(event) {
    event.preventDefault();
    
    const colleagueId = document.getElementById('colleague').value;
    const startDate = document.getElementById('startDate').value;
    const startTime = document.getElementById('startTime').value;
    const endDate = document.getElementById('endDate').value;
    const endTime = document.getElementById('endTime').value;
    
    if (!colleagueId || !startDate || !startTime || !endDate || !endTime) {
        showStatus('error', 'Lütfen tüm alanları doldurun!');
        return;
    }
    
    const colleague = colleagues.find(c => c.id == colleagueId);
    
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
            .replaceAll('{colleagueName}', colleague.name)
            .replaceAll('{email}', colleague.email)
            .replaceAll('{phone}', colleague.phone)
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

// Get user profile information
function getUserProfile() {
    return new Promise((resolve, reject) => {
        console.log('getUserProfile:' + Office.context.mailbox.userProfile.displayName);
        if (typeof Office !== 'undefined' && Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) {
            // userProfile is a direct property, not a method
            const userProfile = Office.context.mailbox.userProfile;
            resolve({
                displayName: userProfile.displayName || 'Kullanıcı',
                emailAddress: userProfile.emailAddress || 'user@oztiryakiler.com.tr',
                jobTitle: userProfile.jobTitle || 'Pozisyon'
            });
        } else {
            // Fallback for testing or when Office context is not available
            resolve({
                displayName: 'Test Kullanıcısı',
                emailAddress: 'test@oztiryakiler.com.tr',
                jobTitle: 'Test Pozisyonu'
            });
        }
    });
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
function setOOFViaEws(startLocal, endLocal, internalMsg, externalMsg, audience="All") {
  const email = Office.context.mailbox.userProfile.emailAddress;
  const start = toLocalNaive(startLocal);
  const end   = toLocalNaive(endLocal);
  // CDATA kullanıyorsan ]]>
  const safeInternal = internalMsg.includes("]]>") ? xmlEscape(internalMsg) : `<![CDATA[${internalMsg}]]>`;
  const safeExternal = externalMsg.includes("]]>") ? xmlEscape(externalMsg) : `<![CDATA[${externalMsg}]]>`;

  const soap = `
  <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                 xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
    <soap:Header>
      <t:RequestServerVersion Version="Exchange2013"/>
    </soap:Header>
    <soap:Body>
      <SetUserOofSettingsRequest xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
        <Mailbox>${email}</Mailbox>
        <UserOofSettings>
          <OofState>Scheduled</OofState>
          <ExternalAudience>${audience}</ExternalAudience>
          <Duration>
            <StartTime>${start}</StartTime>
            <EndTime>${end}</EndTime>
          </Duration>
          <InternalReply><Message>${safeInternal}</Message></InternalReply>
          <ExternalReply><Message>${safeExternal}</Message></ExternalReply>
        </UserOofSettings>
      </SetUserOofSettingsRequest>
    </soap:Body>
  </soap:Envelope>`;

  console.log('EWS SET - Start Time (Local):', start);
  console.log('EWS SET - End Time (Local):', end);

  return new Promise((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(soap, res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(res.error);
      
      console.log('EWS SET Response:', res);
      console.log('EWS SET Response XML:', res.value);
      
      // Parse XML response more thoroughly
      const xml = new window.DOMParser().parseFromString(res.value, "text/xml");
      
      // Try multiple possible response code locations
      let rc = xml.getElementsByTagName("m:ResponseCode")[0] 
            || xml.getElementsByTagName("ResponseCode")[0]
            || xml.querySelector("ResponseCode")
            || xml.querySelector("m\\:ResponseCode");
      
      // Also check for error elements
      const errorCode = xml.getElementsByTagName("ErrorCode")[0] 
                     || xml.getElementsByTagName("m:ErrorCode")[0];
      const faultString = xml.getElementsByTagName("faultstring")[0];
      
      console.log('EWS SET Response Code:', rc ? rc.textContent : 'Not found');
      console.log('EWS SET Error Code:', errorCode ? errorCode.textContent : 'None');
      console.log('EWS SET Fault String:', faultString ? faultString.textContent : 'None');
      
      // Check for success conditions
      if (rc && rc.textContent === "NoError") {
        console.log('EWS SET successful, verifying with GET...');
        resolve(res.value);
      } else if (!rc && !errorCode && !faultString) {
        // No explicit error, might be successful
        console.log('EWS SET - No error codes found, assuming success');
        resolve(res.value);
      } else {
        const errorMsg = errorCode ? errorCode.textContent : 
                        faultString ? faultString.textContent :
                        rc ? rc.textContent : 'Unknown error';
        reject(new Error(`EWS ResponseCode: ${errorMsg}`));
      }
    });
  });
}

// GET function to verify OOF settings
function getOOFViaEws() {
  const email = Office.context.mailbox.userProfile.emailAddress;
  const soap = `
  <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Body>
      <GetUserOofSettingsRequest xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
        <Mailbox xmlns="http://schemas.microsoft.com/exchange/services/2006/types">${email}</Mailbox>
      </GetUserOofSettingsRequest>
    </soap:Body>
  </soap:Envelope>`;
  
  return new Promise((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(soap, res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(res.error);
      const xml = new window.DOMParser().parseFromString(res.value, "text/xml");
      const state = xml.getElementsByTagName("OofState")[0]?.textContent;
      const start = xml.getElementsByTagName("StartTime")[0]?.textContent;
      const end   = xml.getElementsByTagName("EndTime")[0]?.textContent;
      const ext   = xml.getElementsByTagName("ExternalAudience")[0]?.textContent;
      
      console.log('EWS GET Results:', { state, start, end, ext });
      console.log('EWS GET Response XML:', res.value);
      
      resolve({ state, start, end, ext, raw: res.value });
    });
  });
}

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

