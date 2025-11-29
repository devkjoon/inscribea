/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.addEventListener('DOMContentLoaded', initialize);
  }
});

let emailContext = null;

function initialize() {
  // Load current email context
  loadEmailContext();

  // Event listeners
  document.getElementById('generate-btn').addEventListener('click', generateEmail);
  document.getElementById('use-email-context').addEventListener('click', loadEmailContext);
  document.getElementById('insert-btn').addEventListener('click', insertEmail);
  document.getElementById('copy-btn').addEventListener('click', copyToClipboard);
}

async function loadEmailContext() {
  try {
    const item = Office.context.mailbox.item;
    
    if (!item) {
      showMessage('No email item selected', 'error');
      return;
    }

    // Get email properties
    const subject = item.subject;
    const from = item.from ? item.from.emailAddress : 'Unknown';
    
    // Get body as plain text
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        emailContext = {
          subject: subject,
          from: from,
          body: result.value
        };

        // Display context info
        displayEmailContext(emailContext);
        showMessage('Email context loaded successfully', 'success');
      } else {
        console.error('Error loading email body:', result.error);
        emailContext = {
          subject: subject,
          from: from,
          body: ''
        };
        displayEmailContext(emailContext);
      }
    });

  } catch (error) {
    console.error('Error loading email context:', error);
    showMessage('Could not load email context', 'error');
  }
}

function displayEmailContext(context) {
  const contextInfo = document.getElementById('email-context-info');
  const contextDetails = document.getElementById('context-details');
  
  if (context) {
    contextInfo.style.display = 'block';
    contextDetails.innerHTML = `
      <p><strong>Subject:</strong> ${escapeHtml(context.subject || 'No subject')}</p>
      <p><strong>From:</strong> ${escapeHtml(context.from || 'Unknown')}</p>
      <p><strong>Preview:</strong> ${escapeHtml((context.body || '').substring(0, 200))}${context.body && context.body.length > 200 ? '...' : ''}</p>
    `;
  } else {
    contextInfo.style.display = 'none';
  }
}

async function generateEmail() {
  const promptInput = document.getElementById('prompt-input');
  const generateBtn = document.getElementById('generate-btn');
  const generateText = document.getElementById('generate-text');
  const generateSpinner = document.getElementById('generate-spinner');
  const generatedEmail = document.getElementById('generated-email');
  const insertBtn = document.getElementById('insert-btn');
  const copyBtn = document.getElementById('copy-btn');

  const prompt = promptInput.value.trim();

  if (!prompt) {
    showMessage('Please enter a prompt', 'error');
    return;
  }

  // Show loading state
  generateBtn.disabled = true;
  generateText.style.display = 'none';
  generateSpinner.style.display = 'inline-block';
  hideMessages();

  try {
    const response = await fetch('https://localhost:3000/api/generate-email', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        prompt: prompt,
        emailContext: emailContext
      })
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.message || 'Failed to generate email');
    }

    const data = await response.json();
    
    // Display generated email
    generatedEmail.value = data.email;
    insertBtn.disabled = false;
    copyBtn.disabled = false;
    
    showMessage('Email generated successfully!', 'success');

  } catch (error) {
    console.error('Error generating email:', error);
    showMessage(`Error: ${error.message}. Make sure the server is running and OPENAI_API_KEY is configured.`, 'error');
    generatedEmail.value = '';
    insertBtn.disabled = true;
    copyBtn.disabled = true;
  } finally {
    // Reset loading state
    generateBtn.disabled = false;
    generateText.style.display = 'inline';
    generateSpinner.style.display = 'none';
  }
}

function insertEmail() {
  const generatedEmail = document.getElementById('generated-email');
  const emailText = generatedEmail.value.trim();

  if (!emailText) {
    showMessage('No email to insert', 'error');
    return;
  }

  const item = Office.context.mailbox.item;
  
  if (!item) {
    showMessage('No email item selected', 'error');
    return;
  }

  // Insert into the body of the current email
  item.body.setSelectedDataAsync(
    emailText,
    { coercionType: Office.CoercionType.Html },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        showMessage('Email inserted successfully!', 'success');
      } else {
        console.error('Error inserting email:', result.error);
        showMessage('Failed to insert email', 'error');
      }
    }
  );
}

async function copyToClipboard() {
  const generatedEmail = document.getElementById('generated-email');
  const emailText = generatedEmail.value.trim();

  if (!emailText) {
    showMessage('No email to copy', 'error');
    return;
  }

  try {
    await navigator.clipboard.writeText(emailText);
    showMessage('Copied to clipboard!', 'success');
  } catch (error) {
    console.error('Error copying to clipboard:', error);
    // Fallback for older browsers
    const textArea = document.createElement('textarea');
    textArea.value = emailText;
    textArea.style.position = 'fixed';
    textArea.style.opacity = '0';
    document.body.appendChild(textArea);
    textArea.select();
    try {
      document.execCommand('copy');
      showMessage('Copied to clipboard!', 'success');
    } catch (err) {
      showMessage('Failed to copy to clipboard', 'error');
    }
    document.body.removeChild(textArea);
  }
}

function showMessage(message, type) {
  const errorMsg = document.getElementById('error-message');
  const successMsg = document.getElementById('success-message');

  if (type === 'error') {
    errorMsg.textContent = message;
    errorMsg.style.display = 'block';
    successMsg.style.display = 'none';
  } else {
    successMsg.textContent = message;
    successMsg.style.display = 'block';
    errorMsg.style.display = 'none';
  }

  // Auto-hide after 5 seconds
  setTimeout(hideMessages, 5000);
}

function hideMessages() {
  document.getElementById('error-message').style.display = 'none';
  document.getElementById('success-message').style.display = 'none';
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

