const { adminPassword, adminUsername, setSessionCookie } = require("./_auth");

module.exports = function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const { username, password } = req.body || {};
  if (username === adminUsername() && password === adminPassword()) {
    setSessionCookie(res, username);
    res.status(200).json({ success: true });
    return;
  }

  res.status(401).json({ success: false, error: "Invalid username or password" });
};
