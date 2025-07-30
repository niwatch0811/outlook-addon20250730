Office.onReady(() => {
  // Office.jsが完全にロードされた後に実行
  document.getElementById("checkBtn").onclick = checkRecipients;
});

function checkRecipients() {
  const item = Office.context.mailbox.item;

  let resultText = "";

  const groups = [
    { name: "To", list: item.to || [] },
    { name: "Cc", list: item.cc || [] },
    { name: "Bcc", list: item.bcc || [] },
  ];

  groups.forEach(group => {
    resultText += `[${group.name}]\n`;
    group.list.forEach(recipient => {
      const email = (recipient.emailAddress || "").toLowerCase();
      const display = recipient.displayName || "";
      const domain = email.split("@")[1] || "";
      const externalMark = domain !== "tanida.co.jp" ? "【社外】" : "";
      resultText += `${externalMark}${display} <${email}>\n`;
    });
    resultText += "\n";
  });

  document.getElementById("result").textContent = resultText;
}
