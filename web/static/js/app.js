const API = "http://localhost:8000";

async function start() {
  const file = document.getElementById("fileInput").files[0];
  if (!file) return alert("ファイルを選択してください");

  const formData = new FormData();
  formData.append("file", file);

  document.getElementById("status").innerText = "アップロード中...";

  const res = await fetch(`${API}/api/transcribe`, {
    method: "POST",
    body: formData
  });
  const data = await res.json();

  poll(data.job_id);
}

async function poll(jobId) {
  const timer = setInterval(async () => {
    const res = await fetch(`${API}/api/status/${jobId}`);
    const data = await res.json();

    document.getElementById("status").innerText =
      `処理中 ${data.progress || 0}%`;

    if (data.status === "completed") {
      clearInterval(timer);
      const r = await fetch(`${API}/api/result/${jobId}`);
      const result = await r.json();
      document.getElementById("result").value = result.text;
      document.getElementById("status").innerText = "完了 🎉";
    }
  }, 2000);
}