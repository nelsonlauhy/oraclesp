console.log('Custom Menu Script Loaded');

// Example: Add a custom button to SharePoint command bar
setTimeout(() => {
  const toolbar = document.querySelector('[data-automationid="CommandBar"]');
  if (toolbar) {
    const btn = document.createElement('button');
    btn.innerText = 'Site Link';
    btn.style.marginLeft = '10px';
    btn.style.padding = '5px 10px';
    btn.style.backgroundColor = '#0078d4';
    btn.style.color = '#ffffff';
    btn.style.border = 'none';
    btn.style.borderRadius = '4px';
    btn.style.cursor = 'pointer';

    btn.onclick = () => {
      alert('https://oraclegrouprealty.sharepoint.com/sites/OracleGroupPortal');
    };

    toolbar.appendChild(btn);
  }
}, 3000);
