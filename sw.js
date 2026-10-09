/**
 * NexChat — Google Apps Script Backend
 *
 * SETUP:
 * 1. Create a Google Spreadsheet.
 * 2. Extensions → Apps Script.
 * 3. Paste this code into Code.gs.
 * 4. Save the project.
 * 5. Run setupDatabase().
 * 6. Run setupDriveFolders().
 * 7. Deploy as Web app.
 */

const CFG = {
  SPREADSHEET_ID: "",
  ROOT_FOLDER: "NexChat_Data",
  INITIAL_ADMIN_USERNAME: "admin",
  SESSION_TTL_HOURS: 24 * 30,
  PASSWORD_ITERATIONS: 4000,
  MAX_UPLOAD_BYTES: 10 * 1024 * 1024,
  APP_NAME: "NexChat"
};

const SHEETS = {
  Users: [
    "UserID", "Username", "PasswordHash", "PasswordSalt",
    "FullName", "Email", "ProfilePhotoFileID", "ProfilePhotoURL",
    "About", "CreatedAt", "LastSeen", "IsActive", "IsAdmin"
  ],
  Sessions: ["Token", "UserID", "CreatedAt", "ExpiresAt"],
  FriendRequests: [
    "RequestID", "FromUserID", "ToUserID", "Status",
    "CreatedAt", "UpdatedAt"
  ],
  Friendships: ["FriendshipID", "UserA", "UserB", "CreatedAt"],
  Conversations: ["ConversationID", "CreatedAt", "LastActivity"],
  ConversationMembers: ["RowID", "ConversationID", "UserID", "JoinedAt"],
  Messages: [
    "MessageID", "ConversationID", "SenderID", "MessageType",
    "MessageText", "FileID", "FileName", "FileURL", "FileSize",
    "ReplyToMessageID", "Timestamp", "EditedAt",
    "DeletedForEveryone", "ReadBy", "HiddenFor",
    "ReactionsJSON", "ClientMessageID"
  ]
};

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "health";

  if (action === "health") {
    return jsonOut({
      success: true,
      status: "online",
      app: CFG.APP_NAME,
      time: new Date().toISOString()
    });
  }

  return jsonOut({
    success: false,
    message: "Use POST for API calls"
  });
}

function doPost(e) {
  try {
    let body = {};

    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    const action = body.action;

    if (!action) {
      return jsonOut({
        success: false,
        message: "Missing action"
      });
    }

    const handler = API[action];

    if (typeof handler !== "function") {
      return jsonOut({
        success: false,
        message: "Unknown action: " + action
      });
    }

    const publicActions = [
      "register",
      "login",
      "checkUsername"
    ];

    if (!publicActions.includes(action)) {
      const user = requireAuth(body.token);

      if (!user) {
        return jsonOut({
          success: false,
          message: "Unauthorized. Please sign in again."
        });
      }

      body.__user = user;
    }

    return jsonOut(handler(body));

  } catch (err) {
    return jsonOut({
      success: false,
      message: "Server error: " + err.message
    });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ==================== DATABASE SETUP ==================== */

function setupDatabase() {
  const ss = getSpreadsheet();

  Object.entries(SHEETS).forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);

    if (!sh) {
      sh = ss.insertSheet(name);
    }

    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#10B981")
        .setFontColor("#ffffff");

      sh.setFrozenRows(1);

    } else {
      // Add missing columns without deleting existing data.
      const existing = sh
        .getRange(1, 1, 1, sh.getLastColumn())
        .getValues()[0]
        .map(String);

      headers.forEach(header => {
        if (!existing.includes(header)) {
          const col = sh.getLastColumn() + 1;

          sh.getRange(1, col)
            .setValue(header)
            .setFontWeight("bold")
            .setBackground("#10B981")
            .setFontColor("#ffffff");
        }
      });
    }
  });

  const def = ss.getSheetByName("Sheet1");

  if (def && def.getLastRow() === 0) {
    ss.deleteSheet(def);
  }

  return "Database ready. Sheets: " +
    Object.keys(SHEETS).join(", ");
}

function setupDriveFolders() {
  const root = getRootFolder();

  ["Profiles", "Chat_Files", "Group_Photos"].forEach(name => {
    if (!root.getFoldersByName(name).hasNext()) {
      root.createFolder(name);
    }
  });

  return "Drive folders ready";
}

function createInitialAdmin() {
  const users = getSheet("Users");
  const username = CFG.INITIAL_ADMIN_USERNAME;

  if (findRow(users, "Username", username)) {
    return "Admin already exists";
  }

  const salt = randHex(16);
  const hash = hashPassword("ChangeMe123!" + salt);

  users.appendRow([
    uuid(),
    username,
    hash,
    salt,
    "Administrator",
    "",
    "",
    "",
    "System admin",
    new Date(),
    new Date(),
    true,
    true
  ]);

  return "Admin created: " + username +
    " / ChangeMe123! — change immediately";
}

/* ==================== LOW LEVEL HELPERS ==================== */

function getSpreadsheet() {
  if (CFG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(name);

  if (!sh) {
    setupDatabase();
    sh = ss.getSheetByName(name);
  }

  return sh;
}

function getRootFolder() {
  const folders = DriveApp.getFoldersByName(CFG.ROOT_FOLDER);

  return folders.hasNext()
    ? folders.next()
    : DriveApp.createFolder(CFG.ROOT_FOLDER);
}

function getSubFolder(name) {
  const root = getRootFolder();
  const folders = root.getFoldersByName(name);

  return folders.hasNext()
    ? folders.next()
    : root.createFolder(name);
}

function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0];

  return values.slice(1).map((row, i) => {
    const obj = { __row: i + 2 };

    headers.forEach((header, j) => {
      obj[header] = row[j];
    });

    return obj;
  });
}

function findRow(sheet, header, value) {
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return null;
  }

  const index = values[0].indexOf(header);

  if (index === -1) {
    return null;
  }

  for (let i = 1; i < values.length; i++) {
    if (values[i][index] === value) {
      return {
        row: i + 1,
        data: values[i],
        headers: values[0]
      };
    }
  }

  return null;
}

function findAllRows(sheet, header, value) {
  const values = sheet.getDataRange().getValues();
  const output = [];

  if (values.length < 2) {
    return output;
  }

  const index = values[0].indexOf(header);

  for (let i = 1; i < values.length; i++) {
    if (values[i][index] === value) {
      output.push({
        row: i + 1,
        data: values[i],
        headers: values[0]
      });
    }
  }

  return output;
}

function updateCell(sheet, row, headerName, value) {
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const col = headers.indexOf(headerName) + 1;

  if (col > 0) {
    sheet.getRange(row, col).setValue(value);
  }
}

function uuid() {
  return Utilities.getUuid().replace(/-/g, "").slice(0, 24);
}

function randHex(bytes) {
  const array = new Uint8Array(bytes);

  for (let i = 0; i < bytes; i++) {
    array[i] = Math.floor(Math.random() * 256);
  }

  return Array.from(array)
    .map(x => x.toString(16).padStart(2, "0"))
    .join("");
}

/* ==================== PASSWORD HELPERS ==================== */

function hashPassword(password) {
  let data = password;

  for (let i = 0; i < CFG.PASSWORD_ITERATIONS; i++) {
    data = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      data + i
    )
      .map(b => (b + 256) % 256)
      .map(b => String.fromCharCode(b))
      .join("");
  }

  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      data
    )
  );
}

function verifyPassword(password, salt, hash) {
  return hashPassword(password + salt) === hash;
}

/* ==================== AUTHENTICATION ==================== */

function issueSession(userId) {
  const token = uuid() + uuid();
  const now = new Date();

  const expiry = new Date(
    now.getTime() + CFG.SESSION_TTL_HOURS * 3600 * 1000
  );

  getSheet("Sessions").appendRow([
    token,
    userId,
    now,
    expiry
  ]);

  const user = findRow(getSheet("Users"), "UserID", userId);

  if (user) {
    updateCell(getSheet("Users"), user.row, "LastSeen", now);
  }

  return token;
}

function requireAuth(token) {
  if (!token) {
    return null;
  }

  const session = findRow(getSheet("Sessions"), "Token", token);

  if (!session) {
    return null;
  }

  if (new Date(session.data[3]).getTime() < Date.now()) {
    return null;
  }

  const user = findRow(
    getSheet("Users"),
    "UserID",
    session.data[1]
  );

  if (!user || user.data[11] !== true) {
    return null;
  }

  return rowToUser(user);
}

function rowToUser(user) {
  const headers = user.headers;
  const row = user.data;
  const get = name => row[headers.indexOf(name)];

  return {
    userId: get("UserID"),
    username: get("Username"),
    fullName: get("FullName"),
    email: get("Email"),
    avatarUrl: get("ProfilePhotoURL"),
    about: get("About"),
    createdAt: get("CreatedAt"),
    lastSeen: get("LastSeen"),
    isActive: get("IsActive") === true,
    isAdmin: get("IsAdmin") === true
  };
}

function publicUser(user, isSelf) {
  const base = {
    userId: user.userId,
    username: user.username,
    fullName: user.fullName,
    avatarUrl: user.avatarUrl,
    color: "#" + Math.abs(
      user.username.split("").reduce(
        (hash, char) => (hash << 5) - hash + char.charCodeAt(0),
        0
      )
    ).toString(16).slice(0, 6),
    about: user.about,
    online: (Date.now() - new Date(user.lastSeen).getTime()) < 60000,
    lastSeen: user.lastSeen,
    isAdmin: user.isAdmin
  };

  if (isSelf) {
    base.email = user.email;
  }

  return base;
}

function ok(data, message) {
  return {
    success: true,
    message: message || "OK",
    data: data || {}
  };
}

function err(message) {
  return {
    success: false,
    message: message
  };
}

/* ==================== API ==================== */

const API = {};

function userCanAccessConversation_(conversationId, userId) {
  if (!conversationId) {
    return false;
  }

  const members = sheetToObjects(
    getSheet("ConversationMembers")
  );

  return members.some(member =>
    member.ConversationID === conversationId &&
    member.UserID === userId
  );
}

function messageRowByClientId_(clientMessageId, senderId) {
  if (!clientMessageId) {
    return null;
  }

  const hit = findRow(
    getSheet("Messages"),
    "ClientMessageID",
    String(clientMessageId)
  );

  if (!hit) {
    return null;
  }

  if (hit.data[hit.headers.indexOf("SenderID")] !== senderId) {
    return null;
  }

  return hit;
}

/* ==================== AUTH API ==================== */

API.register = function(b) {
  if (!/^[a-zA-Z0-9_]{3,24}$/.test(b.username || "")) {
    return err(
      "Username must be 3–24 characters (letters, numbers, underscore only)"
    );
  }

  if (!b.password || b.password.length < 6) {
    return err("Password must be at least 6 characters");
  }

  if (!b.fullName || b.fullName.length < 2) {
    return err("Full name is required");
  }

  const users = getSheet("Users");

  if (findRow(users, "Username", b.username)) {
    return err("Username already taken. Try another.");
  }

  const userId = uuid();
  const salt = randHex(16);
  const hash = hashPassword(b.password + salt);
  const isAdmin = b.username === CFG.INITIAL_ADMIN_USERNAME;

  users.appendRow([
    userId,
    b.username,
    hash,
    salt,
    b.fullName,
    b.email || "",
    "",
    "",
    "Hey there! I am using " + CFG.APP_NAME + ".",
    new Date(),
    new Date(),
    true,
    isAdmin
  ]);

  return ok({ userId }, "Account created successfully");
};

API.login = function(b) {
  const user = findRow(
    getSheet("Users"),
    "Username",
    b.username
  );

  if (!user) {
    return err("Invalid username or password");
  }

  const headers = user.headers;
  const row = user.data;

  if (!verifyPassword(
    b.password,
    row[headers.indexOf("PasswordSalt")],
    row[headers.indexOf("PasswordHash")]
  )) {
    return err("Invalid username or password");
  }

  if (row[headers.indexOf("IsActive")] !== true) {
    return err("Account disabled");
  }

  const userData = rowToUser(user);
  const token = issueSession(userData.userId);

  return ok({
    token,
    user: publicUser(userData, true)
  }, "Signed in");
};

API.logout = function(b) {
  const sheet = getSheet("Sessions");
  const session = findRow(sheet, "Token", b.token);

  if (session) {
    sheet.deleteRow(session.row);
  }

  return ok({}, "Signed out");
};

API.me = function(b) {
  return ok({
    user: publicUser(b.__user, true)
  });
};

API.checkUsername = function(b) {
  if (!b.username) {
    return ok({ available: false });
  }

  const taken = !!findRow(
    getSheet("Users"),
    "Username",
    b.username
  );

  return ok({ available: !taken });
};

API.heartbeat = function(b) {
  const user = findRow(
    getSheet("Users"),
    "UserID",
    b.__user.userId
  );

  if (user) {
    updateCell(
      getSheet("Users"),
      user.row,
      "LastSeen",
      new Date()
    );
  }

  return ok();
};

/* ==================== USER API ==================== */

API.searchUsers = function(b) {
  const me = b.__user.userId;
  const query = (b.query || "").toLowerCase().replace(/^@/, "");

  if (!query) {
    return ok({ users: [] });
  }

  const users = sheetToObjects(getSheet("Users"))
    .filter(user => user.IsActive === true && user.UserID !== me)
    .filter(user =>
      String(user.Username || "").toLowerCase().includes(query) ||
      String(user.FullName || "").toLowerCase().includes(query)
    )
    .slice(0, 30);

  const friendships = sheetToObjects(getSheet("Friendships"));
  const requests = sheetToObjects(getSheet("FriendRequests"));

  const list = users.map(userRow => {
    const user = rowToUser({
      data: Object.values(userRow).slice(1),
      headers: Object.keys(userRow).slice(1)
    });

    const publicData = publicUser(user, false);
    let relation = "none";

    if (friendships.some(friend =>
      (friend.UserA === me && friend.UserB === userRow.UserID) ||
      (friend.UserA === userRow.UserID && friend.UserB === me)
    )) {
      relation = "friends";

    } else if (requests.some(request =>
      request.FromUserID === me &&
      request.ToUserID === userRow.UserID &&
      request.Status === "pending"
    )) {
      relation = "pending_out";

    } else if (requests.some(request =>
      request.FromUserID === userRow.UserID &&
      request.ToUserID === me &&
      request.Status === "pending"
    )) {
      relation = "pending_in";
    }

    publicData.relation = relation;
    return publicData;
  });

  return ok({ users: list });
};

API.updateProfile = function(b) {
  const users = getSheet("Users");
  const user = findRow(users, "UserID", b.__user.userId);

  if (!user) {
    return err("User not found");
  }

  if (typeof b.fullName === "string" && b.fullName.trim()) {
    updateCell(users, user.row, "FullName", b.fullName.trim());
  }

  if (typeof b.about === "string") {
    updateCell(users, user.row, "About", b.about.trim());
  }

  if (typeof b.avatarUrl === "string") {
    updateCell(users, user.row, "ProfilePhotoURL", b.avatarUrl);
  }

  const fresh = rowToUser(
    findRow(users, "UserID", b.__user.userId)
  );

  return ok({
    user: publicUser(fresh, true)
  }, "Profile updated");
};

/* ==================== FRIEND REQUESTS ==================== */

API.sendRequest = function(b) {
  const me = b.__user.userId;
  const recipient = b.toUserId;

  if (!recipient || recipient === me) {
    return err("Invalid recipient");
  }

  if (!findRow(getSheet("Users"), "UserID", recipient)) {
    return err("User not found");
  }

  const friendships = sheetToObjects(getSheet("Friendships"));

  if (friendships.some(friend =>
    (friend.UserA === me && friend.UserB === recipient) ||
    (friend.UserA === recipient && friend.UserB === me)
  )) {
    return err("You are already friends");
  }

  const requests = sheetToObjects(getSheet("FriendRequests"));

  const existing = requests.find(request =>
    request.FromUserID === me &&
    request.ToUserID === recipient &&
    request.Status === "pending"
  );

  if (existing) {
    return err("Request already sent");
  }

  const reverse = requests.find(request =>
    request.FromUserID === recipient &&
    request.ToUserID === me &&
    request.Status === "pending"
  );

  if (reverse) {
    return err(
      "This user already sent you a request — check your Requests tab"
    );
  }

  getSheet("FriendRequests").appendRow([
    uuid(),
    me,
    recipient,
    "pending",
    new Date(),
    new Date()
  ]);

  return ok({}, "Friend request sent");
};

API.getPendingRequests = function(b) {
  const me = b.__user.userId;

  const requests = sheetToObjects(getSheet("FriendRequests"))
    .filter(request =>
      request.ToUserID === me &&
      request.Status === "pending"
    );

  const users = sheetToObjects(getSheet("Users"));

  const list = requests.map(request => {
    const userRow = users.find(
      user => user.UserID === request.FromUserID
    );

    if (!userRow) {
      return null;
    }

    const user = rowToUser({
      data: Object.values(userRow).slice(1),
      headers: Object.keys(userRow).slice(1)
    });

    const publicData = publicUser(user, false);
    publicData.requestId = request.RequestID;
    publicData.createdAt = request.CreatedAt;

    return publicData;
  }).filter(Boolean);

  return ok({ requests: list });
};

API.acceptRequest = function(b) {
  const me = b.__user.userId;
  const request = findRow(
    getSheet("FriendRequests"),
    "RequestID",
    b.requestId
  );

  if (!request) {
    return err("Request not found");
  }

  const headers = request.headers;

  if (request.data[headers.indexOf("ToUserID")] !== me) {
    return err("Not your request");
  }

  updateCell(
    getSheet("FriendRequests"),
    request.row,
    "Status",
    "accepted"
  );

  updateCell(
    getSheet("FriendRequests"),
    request.row,
    "UpdatedAt",
    new Date()
  );

  const sender = request.data[headers.indexOf("FromUserID")];

  getSheet("Friendships").appendRow([
    uuid(),
    sender,
    me,
    new Date()
  ]);

  return ok({}, "Request accepted");
};

API.declineRequest = function(b) {
  const me = b.__user.userId;
  const request = findRow(
    getSheet("FriendRequests"),
    "RequestID",
    b.requestId
  );

  if (!request) {
    return err("Request not found");
  }

  const headers = request.headers;

  if (request.data[headers.indexOf("ToUserID")] !== me) {
    return err("Not your request");
  }

  updateCell(
    getSheet("FriendRequests"),
    request.row,
    "Status",
    "declined"
  );

  updateCell(
    getSheet("FriendRequests"),
    request.row,
    "UpdatedAt",
    new Date()
  );

  return ok({}, "Request declined");
};

/* ==================== CONVERSATIONS ==================== */

API.getConversations = function(b) {
  const me = b.__user.userId;

  const friendships = sheetToObjects(getSheet("Friendships"))
    .filter(friend =>
      friend.UserA === me || friend.UserB === me
    );

  const conversations = sheetToObjects(getSheet("Conversations"));
  const members = sheetToObjects(getSheet("ConversationMembers"));
  const users = sheetToObjects(getSheet("Users"));
  const messages = sheetToObjects(getSheet("Messages"));

  const list = friendships.map(friend => {
    const otherId = friend.UserA === me
      ? friend.UserB
      : friend.UserA;

    let conversation = conversations.find(item => {
      const conversationMembers = members.filter(
        member => member.ConversationID === item.ConversationID
      );

      return conversationMembers.length === 2 &&
        conversationMembers.some(member => member.UserID === me) &&
        conversationMembers.some(member => member.UserID === otherId);
    });

    const conversationId = conversation
      ? conversation.ConversationID
      : "virt_" + [me, otherId].sort().join("_");

    const peerRow = users.find(user => user.UserID === otherId);

    if (!peerRow) {
      return null;
    }

    const peer = rowToUser({
      data: Object.values(peerRow).slice(1),
      headers: Object.keys(peerRow).slice(1)
    });

    const conversationMessages = messages
      .filter(message =>
        message.ConversationID === conversationId &&
        !String(message.HiddenFor || "").split(",").includes(me)
      )
      .sort((a, b) =>
        new Date(b.Timestamp) - new Date(a.Timestamp)
      );

    const last = conversationMessages[0];
    let lastText = "";

    if (last) {
      if (last.DeletedForEveryone) {
        lastText = "This message was deleted";
      } else if (last.MessageType === "image") {
        lastText = "📷 Photo";
      } else if (last.MessageType === "video") {
        lastText = "🎥 Video";
      } else if (last.MessageType === "file") {
        lastText = "📎 " + (last.FileName || "File");
      } else {
        lastText = last.MessageText || "";
      }
    }

    const unreadCount = conversationMessages.filter(message =>
      message.SenderID !== me &&
      !String(message.ReadBy || "").split(",").includes(me)
    ).length;

    return {
      conversationId,
      type: "direct",
      name: peer.fullName,
      avatarUrl: peer.avatarUrl,
      avatarColor: "#" + Math.abs(
        peer.username.split("").reduce(
          (hash, char) => (hash << 5) - hash + char.charCodeAt(0),
          0
        )
      ).toString(16).slice(0, 6),
      peerUserId: peer.userId,
      peerUsername: peer.username,
      online: (Date.now() - new Date(peer.lastSeen).getTime()) < 60000,
      lastSeen: peer.lastSeen,
      lastMessage: lastText.slice(0, 60),
      lastMessageId: last ? last.MessageID : "",
      lastMessageSenderId: last ? last.SenderID : null,
      lastMessageStatus: last
        ? (
          String(last.ReadBy || "").split(",").includes(otherId)
            ? "read"
            : "sent"
        )
        : null,
      lastActivity: last ? last.Timestamp : friend.CreatedAt,
      unreadCount
    };
  }).filter(Boolean);

  list.sort((a, b) =>
    new Date(b.lastActivity) - new Date(a.lastActivity)
  );

  return ok({ conversations: list });
};

API.createConversation = function(b) {
  const me = b.__user.userId;
  const peer = b.peerUserId;

  if (!peer) {
    return err("Invalid peer");
  }

  const friendships = sheetToObjects(getSheet("Friendships"));

  const isFriend = friendships.some(friend =>
    (friend.UserA === me && friend.UserB === peer) ||
    (friend.UserA === peer && friend.UserB === me)
  );

  if (!isFriend) {
    return err("You must be friends first");
  }

  const members = sheetToObjects(getSheet("ConversationMembers"));
  const conversations = sheetToObjects(getSheet("Conversations"));

  for (const conversation of conversations) {
    const conversationMembers = members.filter(
      member => member.ConversationID === conversation.ConversationID
    );

    if (
      conversationMembers.length === 2 &&
      conversationMembers.some(member => member.UserID === me) &&
      conversationMembers.some(member => member.UserID === peer)
    ) {
      return ok({
        conversationId: conversation.ConversationID
      });
    }
  }

  const conversationId = uuid();
  const now = new Date();

  getSheet("Conversations").appendRow([
    conversationId,
    now,
    now
  ]);

  getSheet("ConversationMembers").appendRow([
    uuid(),
    conversationId,
    me,
    now
  ]);

  getSheet("ConversationMembers").appendRow([
    uuid(),
    conversationId,
    peer,
    now
  ]);

  return ok({ conversationId });
};

/* ==================== MESSAGES ==================== */

API.getMessages = function(b) {
  const me = b.__user.userId;
  const conversationId = b.conversationId;
  const limit = Math.min(b.limit || 100, 300);

  const all = sheetToObjects(getSheet("Messages"))
    .filter(message =>
      message.ConversationID === conversationId
    )
    .filter(message =>
      !String(message.HiddenFor || "").split(",").includes(me)
    );

  all.sort((a, b) =>
    new Date(a.Timestamp) - new Date(b.Timestamp)
  );

  const sliced = all.slice(-limit);
  const users = sheetToObjects(getSheet("Users"));

  const enriched = sliced.map(message => {
    const sender = users.find(
      user => user.UserID === message.SenderID
    );

    const senderName = sender ? sender.FullName : "";
    let replyTo = null;

    if (message.ReplyToMessageID) {
      const repliedMessage = all.find(
        item => item.MessageID === message.ReplyToMessageID
      );

      if (repliedMessage) {
        const replySender = users.find(
          user => user.UserID === repliedMessage.SenderID
        );

        replyTo = {
          messageId: repliedMessage.MessageID,
          senderName: replySender ? replySender.FullName : "",
          preview: (
            repliedMessage.DeletedForEveryone
              ? "deleted"
              : (repliedMessage.MessageText || "📎")
          ).slice(0, 80)
        };
      }
    }

    let reactions = {};

    try {
      reactions = JSON.parse(message.ReactionsJSON || "{}");
    } catch (error) {
      reactions = {};
    }

    let status = "sent";

    if (message.SenderID === me) {
      const readBy = String(message.ReadBy || "")
        .split(",")
        .filter(Boolean);

      if (readBy.length > 0) {
        status = "read";
      } else {
        status = "delivered";
      }
    }

    const deleted =
      message.DeletedForEveryone === true ||
      String(message.DeletedForEveryone).toLowerCase() === "true";

    return {
      messageId: message.MessageID,
      clientMessageId: message.ClientMessageID || "",
      conversationId: message.ConversationID,
      senderId: message.SenderID,
      senderName,
      messageType: deleted ? "text" : message.MessageType,
      text: deleted ? "" : message.MessageText,
      fileUrl: deleted ? "" : message.FileURL,
      fileName: deleted ? "" : message.FileName,
      fileSize: deleted ? 0 : message.FileSize,
      timestamp: message.Timestamp,
      editedAt: message.EditedAt || null,
      deletedForEveryone: deleted,
      status,
      replyTo,
      reactions: deleted ? {} : reactions
    };
  });

  return ok({ messages: enriched });
};

API.sendMessage = function(b) {
  const me = b.__user.userId;
  const lock = LockService.getScriptLock();

  lock.waitLock(15000);

  try {
    // Prevent duplicate sends when a client retries the same message.
    const duplicate = messageRowByClientId_(
      b.clientMessageId,
      me
    );

    if (duplicate) {
      const headers = duplicate.headers;
      const row = duplicate.data;
      const userRow = findRow(getSheet("Users"), "UserID", me);

      return ok({
        message: {
          messageId: row[headers.indexOf("MessageID")],
          clientMessageId: row[headers.indexOf("ClientMessageID")] || "",
          conversationId: row[headers.indexOf("ConversationID")],
          senderId: me,
          senderName: userRow
            ? userRow.data[userRow.headers.indexOf("FullName")]
            : "",
          messageType: row[headers.indexOf("MessageType")],
          text: row[headers.indexOf("MessageText")],
          fileUrl: row[headers.indexOf("FileURL")],
          fileName: row[headers.indexOf("FileName")],
          fileSize: row[headers.indexOf("FileSize")],
          timestamp: row[headers.indexOf("Timestamp")],
          status: "sent",
          reactions: {},
          replyTo: null,
          deletedForEveryone:
            row[headers.indexOf("DeletedForEveryone")] === true
        }
      }, "Already sent");
    }

    let conversationId = b.conversationId;

    if (conversationId && conversationId.startsWith("virt_")) {
      const parts = conversationId.replace("virt_", "").split("_");
      const peer = parts.find(id => id !== me);

      const createResult = API.createConversation({
        __user: b.__user,
        peerUserId: peer
      });

      if (createResult.success) {
        conversationId = createResult.data.conversationId;
      } else {
        return err("Cannot create conversation");
      }
    }

    if (!conversationId) {
      return err("conversationId required");
    }

    if (!userCanAccessConversation_(conversationId, me)) {
      return err("You are not a member of this conversation");
    }

    const type = b.messageType || "text";

    if (type === "text" && (!b.text || !String(b.text).trim())) {
      return err("Message text required");
    }

    const messageId = uuid();
    const now = new Date();

    getSheet("Messages").appendRow([
      messageId,
      conversationId,
      me,
      type,
      b.text || "",
      b.fileId || "",
      b.fileName || "",
      b.fileUrl || "",
      b.fileSize || 0,
      b.replyToMessageId || "",
      now,
      "",
      false,
      "",
      "",
      "{}",
      b.clientMessageId ? String(b.clientMessageId) : ""
    ]);

    const conversation = findRow(
      getSheet("Conversations"),
      "ConversationID",
      conversationId
    );

    if (conversation) {
      updateCell(
        getSheet("Conversations"),
        conversation.row,
        "LastActivity",
        now
      );
    }

    const sender = findRow(getSheet("Users"), "UserID", me);

    return ok({
      message: {
        messageId,
        clientMessageId: b.clientMessageId || "",
        conversationId,
        senderId: me,
        senderName: sender
          ? sender.data[sender.headers.indexOf("FullName")]
          : "",
        messageType: type,
        text: b.text || "",
        fileUrl: b.fileUrl || "",
        fileName: b.fileName || "",
        fileSize: b.fileSize || 0,
        timestamp: now,
        status: "sent",
        reactions: {},
        replyTo: null,
        deletedForEveryone: false
      }
    }, "Message sent");

  } finally {
    lock.releaseLock();
  }
};

API.markRead = function(b) {
  const me = b.__user.userId;

  const message = findRow(
    getSheet("Messages"),
    "MessageID",
    b.upToMessageId
  );

  if (!message) {
    return ok();
  }

  const headers = message.headers;

  const readBy = String(
    message.data[headers.indexOf("ReadBy")] || ""
  ).split(",").filter(Boolean);

  if (!readBy.includes(me)) {
    readBy.push(me);
  }

  updateCell(
    getSheet("Messages"),
    message.row,
    "ReadBy",
    readBy.join(",")
  );

  return ok();
};

API.editMessage = function(b) {
  const message = findRow(
    getSheet("Messages"),
    "MessageID",
    b.messageId
  );

  if (
    !message ||
    message.data[message.headers.indexOf("SenderID")] !== b.__user.userId
  ) {
    return err("Cannot edit");
  }

  updateCell(
    getSheet("Messages"),
    message.row,
    "MessageText",
    b.text
  );

  updateCell(
    getSheet("Messages"),
    message.row,
    "EditedAt",
    new Date()
  );

  return ok({}, "Edited");
};

API.deleteMessage = function(b) {
  const me = b.__user.userId;

  const message = findRow(
    getSheet("Messages"),
    "MessageID",
    b.messageId
  );

  if (!message) {
    return err("Not found");
  }

  const headers = message.headers;
  const conversationId =
    message.data[headers.indexOf("ConversationID")];

  const senderId =
    message.data[headers.indexOf("SenderID")];

  if (!userCanAccessConversation_(conversationId, me)) {
    return err("Not a member of this conversation");
  }

  if (b.scope === "everyone") {
    if (senderId !== me) {
      return err("Only the sender can delete this message for everyone");
    }

    updateCell(getSheet("Messages"), message.row, "DeletedForEveryone", true);
    updateCell(getSheet("Messages"), message.row, "MessageText", "");
    updateCell(getSheet("Messages"), message.row, "FileURL", "");
    updateCell(getSheet("Messages"), message.row, "FileID", "");
    updateCell(getSheet("Messages"), message.row, "FileName", "");
    updateCell(getSheet("Messages"), message.row, "FileSize", 0);
    updateCell(getSheet("Messages"), message.row, "ReactionsJSON", "{}");

  } else if (b.scope === "me" || !b.scope) {
    const hiddenFor = String(
      message.data[headers.indexOf("HiddenFor")] || ""
    ).split(",").filter(Boolean);

    if (!hiddenFor.includes(me)) {
      hiddenFor.push(me);
    }

    updateCell(
      getSheet("Messages"),
      message.row,
      "HiddenFor",
      hiddenFor.join(",")
    );

  } else {
    return err("Invalid delete scope");
  }

  return ok({}, "Deleted");
};

API.addReaction = function(b) {
  const message = findRow(
    getSheet("Messages"),
    "MessageID",
    b.messageId
  );

  if (!message) {
    return err("Not found");
  }

  const headers = message.headers;
  let reactions = {};

  try {
    reactions = JSON.parse(
      message.data[headers.indexOf("ReactionsJSON")] || "{}"
    );
  } catch (error) {
    reactions = {};
  }

  if (!reactions[b.reaction]) {
    reactions[b.reaction] = [];
  }

  if (!reactions[b.reaction].includes(b.__user.userId)) {
    reactions[b.reaction].push(b.__user.userId);
  }

  updateCell(
    getSheet("Messages"),
    message.row,
    "ReactionsJSON",
    JSON.stringify(reactions)
  );

  return ok();
};

API.removeReaction = function(b) {
  const message = findRow(
    getSheet("Messages"),
    "MessageID",
    b.messageId
  );

  if (!message) {
    return err("Not found");
  }

  const headers = message.headers;
  const conversationId =
    message.data[headers.indexOf("ConversationID")];

  if (!userCanAccessConversation_(conversationId, b.__user.userId)) {
    return err("Not a member of this conversation");
  }

  let reactions = {};

  try {
    reactions = JSON.parse(
      message.data[headers.indexOf("ReactionsJSON")] || "{}"
    );
  } catch (error) {
    reactions = {};
  }

  const userId = b.__user.userId;

  const emojis = b.reaction
    ? [String(b.reaction)]
    : Object.keys(reactions);

  emojis.forEach(emoji => {
    if (Array.isArray(reactions[emoji])) {
      reactions[emoji] = reactions[emoji].filter(
        id => id !== userId
      );

      if (reactions[emoji].length === 0) {
        delete reactions[emoji];
      }
    }
  });

  updateCell(
    getSheet("Messages"),
    message.row,
    "ReactionsJSON",
    JSON.stringify(reactions)
  );

  return ok({ reactions }, "Reaction removed");
};

API.clearChat = function(b) {
  const me = b.__user.userId;

  if (!userCanAccessConversation_(b.conversationId, me)) {
    return err("Not a member of this conversation");
  }

  const messages = sheetToObjects(getSheet("Messages"))
    .filter(message =>
      message.ConversationID === b.conversationId
    );

  messages.forEach(message => {
    const row = findRow(
      getSheet("Messages"),
      "MessageID",
      message.MessageID
    );

    if (!row) {
      return;
    }

    const headers = row.headers;

    const hiddenFor = String(
      row.data[headers.indexOf("HiddenFor")] || ""
    ).split(",").filter(Boolean);

    if (!hiddenFor.includes(me)) {
      hiddenFor.push(me);
    }

    updateCell(
      getSheet("Messages"),
      row.row,
      "HiddenFor",
      hiddenFor.join(",")
    );
  });

  return ok({}, "Chat cleared");
};

/* ==================== FILE UPLOADS ==================== */

API.uploadFile = function(b) {
  if (!b.base64) {
    return err("No file data");
  }

  const approximateSize = Math.ceil(b.base64.length * 0.75);

  if (approximateSize > CFG.MAX_UPLOAD_BYTES) {
    return err(
      "File too large (max " +
      (CFG.MAX_UPLOAD_BYTES / 1024 / 1024) +
      " MB)"
    );
  }

  const mime = b.mime || "application/octet-stream";

  const folderName =
    mime.startsWith("image/") && b.name === "avatar.jpg"
      ? "Profiles"
      : "Chat_Files";

  const folder = getSubFolder(folderName);
  const bytes = Utilities.base64Decode(b.base64);
  const blob = Utilities.newBlob(
    bytes,
    mime,
    b.name || ("file_" + Date.now())
  );

  const file = folder.createFile(blob);

  try {
    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );
  } catch (error) {
    // Sharing settings may be restricted by the account.
  }

  const fileId = file.getId();
  const fileUrl =
    "https://drive.google.com/uc?export=view&id=" + fileId;

  return ok({
    fileId,
    fileUrl,
    size: approximateSize
  }, "Uploaded");
};

/* ==================== HEALTH CHECK ==================== */

API.health = function() {
  return ok({
    status: "online",
    app: CFG.APP_NAME
  });
};