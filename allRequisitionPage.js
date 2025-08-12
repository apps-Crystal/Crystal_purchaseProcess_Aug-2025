function allRequisitionsPage() {
  return `
  <style>
    .card-grid { display: flex; flex-wrap: wrap; gap: 20px; }
    .requisition-card {
      background: #fff; border-radius: 8px; box-shadow: 0 2px 8px #eee;
      padding: 20px; width: 320px; cursor: pointer; transition: box-shadow 0.2s;
    }
    .requisition-card:hover { box-shadow: 0 4px 16px #ccc; }
    .modal-bg {
      display: none; position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
      background: rgba(0,0,0,0.3); align-items: center; justify-content: center; z-index: 9999;
    }
    .modal-bg.active { display: flex; }
    .modal-content {
      background: #fff; border-radius: 8px; padding: 30px; min-width: 350px; max-width: 90vw;
      box-shadow: 0 8px 32px rgba(0,0,0,0.15);
    }
    .modal-content h3 { margin-top: 0; }
    .close-modal { float: right; cursor: pointer; font-size: 1.3em; color: #888; }
    .modal-details { margin-top: 10px; }
    .modal-details div { margin-bottom: 8px; }
  </style>
  <div>
    <h2>All Requisitions</h2>
    <div id="requisitionCards" class="card-grid"></div>
    <div id="modalBg" class="modal-bg">
      <div class="modal-content">
        <span class="close-modal" onclick="document.getElementById('modalBg').classList.remove('active')">&times;</span>
        <div id="modalDetails"></div>
      </div>
    </div>
  </div>
  <script>
    google.script.run.withSuccessHandler(function(data) {
      const container = document.getElementById('requisitionCards');
      if (!data || !data.length) {
        container.innerHTML = '<p>No requisitions found.</p>';
        return;
      }
      container.innerHTML = data.map((req, idx) => \`
        <div class="requisition-card" onclick="showReqDetails(\${idx})">
          <strong>PR ID:</strong> \${req['Requisition ID'] || ''}<br>
          <strong>Requested By:</strong> \${req['Requested By'] || ''}<br>
          <strong>Site:</strong> \${req['Site'] || ''}<br>
          <strong>Status:</strong> \${req['Current Status'] || ''}<br>
          <strong>Date:</strong> \${req['Date of Requisition'] || ''}
        </div>
      \`).join('');
      window.allReqData = data;
    }).getAllRequisitions();

    window.showReqDetails = function(idx) {
      const req = window.allReqData[idx];
      let html = '<h3>Requisition Details</h3><div class="modal-details">';
      for (const key in req) {
        html += '<div><strong>' + key + ':</strong> ' + (req[key] || '') + '</div>';
      }
      html += '</div>';
      document.getElementById('modalDetails').innerHTML = html;
      document.getElementById('modalBg').classList.add('active');
    }
  </script>
  `;
}