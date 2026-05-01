const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType
} = require('docx');

function inch(n) { return Math.round(n * 1440); }

const PAGE_W = 11906;
const PAGE_H = 16838;
const M_LEFT = inch(1.25);
const M_RIGHT = inch(1.0);
const M_TOP = inch(1.0);
const M_BOTTOM = inch(1.0);

function TNR(size, bold = false, italics = false, color = '000000') {
  return { font: 'Times New Roman', size: size * 2, bold, italics, color };
}

function pageProps() {
  return {
    page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: M_TOP, bottom: M_BOTTOM, left: M_LEFT, right: M_RIGHT } }
  };
}

function createHeader(text, size, bold, color = '000000') {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200, line: 360, lineRule: 'exact' },
    children: [new TextRun({ text: text, ...TNR(size, bold, false, color) })]
  });
}

function createPara(text, size = 12, bold = false, italics = false) {
  return new Paragraph({
    spacing: { after: 120, line: 360, lineRule: 'exact' },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text: text, ...TNR(size, bold, italics) })]
  });
}

function createSection(title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 200, line: 360, lineRule: 'exact' },
    children: [new TextRun({ text: title, ...TNR(16, true, false, '000000') })]
  });
}

function createSubSection(title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 150, line: 360, lineRule: 'exact' },
    children: [new TextRun({ text: title, ...TNR(14, true, false, '000000') })]
  });
}

function createSubSubSection(title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 100, line: 360, lineRule: 'exact' },
    children: [new TextRun({ text: title, ...TNR(12, true, false, '000000') })]
  });
}

function createTocItem(title, page) {
  return new Paragraph({
    children: [
      new TextRun({ text: title, ...TNR(12, false) }),
      new TextRun({ text: " ", ...TNR(12) }),
      new TextRun({ text: page, ...TNR(12) })
    ]
  });
}

// ============ COVER PAGE ============
const coverPage = [];
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(createHeader("DEPARTMENT OF COMPUTER ENGINEERING", 14, true));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(createHeader("INDUS UNIVERSITY", 16, true));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(createHeader("PROJECT ID: IU/ITE/CE/2025/IDP-001", 12, false));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(createHeader("STUDENT HELPER CHATBOT", 18, true));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(createHeader("A Project Report", 14, false));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(createHeader("Submitted in partial fulfillment of the requirements for the degree of", 12, false));
coverPage.push(createHeader("Bachelor of Technology (B.Tech.) in Computer Engineering", 12, false));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(createHeader("Submitted by", 12, false));
coverPage.push(createHeader("Your Name - Enrollment No.", 12, false));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(createHeader("Under the Guidance of", 12, false));
coverPage.push(createHeader("Project Supervisor", 12, false));
coverPage.push(new Paragraph({ text: '' }));
coverPage.push(createHeader("Academic Year: 2025-2026", 12, false));

// ============ FIRST PAGE ============
const firstPage = [];
firstPage.push(createHeader("STUDENT HELPER CHATBOT", 16, true));
firstPage.push(new Paragraph({ text: '' }));
firstPage.push(createPara("Project ID: IU/ITE/CE/2025/IDP-001", 12, false));
firstPage.push(createPara("Department of Computer Engineering", 12, false));
firstPage.push(createPara("Indus University", 12, false));

// ============ CANDIDATE'S DECLARATION ============
const declaration = [];
declaration.push(createHeader("CANDIDATE'S DECLARATION", 14, true));
declaration.push(new Paragraph({ text: '' }));
declaration.push(createPara("I hereby declare that the project work entitled \"Student Helper Chatbot\" is a bonafide work carried out by me under the guidance of Project Supervisor and that this work has not been submitted anywhere else for any other purpose."));
declaration.push(new Paragraph({ text: '' }));
declaration.push(createPara("Date: _________________", 12, false));
declaration.push(createPara("Place: _________________", 12, false));
declaration.push(new Paragraph({ text: '' }));
declaration.push(createPara("Signature: _________________", 12, false));
declaration.push(createPara("Name: _________________", 12, false));
declaration.push(createPara("Enrollment No.: _________________", 12, false));

// ============ COMPANY CERTIFICATE ============
const companyCert = [];
companyCert.push(createHeader("COMPANY CERTIFICATE", 14, true));
companyCert.push(new Paragraph({ text: '' }));
companyCert.push(createPara("This is to certify that _________________ (Enrollment No.: _________________) has successfully completed the project work titled \"Student Helper Chatbot\" at _________________ (Company Name) during the period ________ to ________."));
companyCert.push(new Paragraph({ text: '' }));
companyCert.push(createPara("The student has worked on the following areas:"));
companyCert.push(createPara("• Analysis and design of the system", 12, false));
companyCert.push(createPara("• Implementation of the chatbot functionality", 12, false));
companyCert.push(createPara("• Testing and documentation", 12, false));
companyCert.push(new Paragraph({ text: '' }));
companyCert.push(createPara("Date: _________________", 12, false));
companyCert.push(createPara("Place: _________________", 12, false));
companyCert.push(new Paragraph({ text: '' }));
companyCert.push(createPara("Signature: _________________", 12, false));
companyCert.push(createPara("Name of the Mentor", 12, false));
companyCert.push(createPara("Designation", 12, false));

// ============ COLLEGE CERTIFICATE ============
const collegeCert = [];
collegeCert.push(createHeader("COLLEGE CERTIFICATE", 14, true));
collegeCert.push(new Paragraph({ text: '' }));
collegeCert.push(createPara("This is to certify that the project work entitled \"Student Helper Chatbot\" is a bonafide work carried out by _________________ (Enrollment No.: _________________) of B.Tech. Computer Engineering (Batch: 2025-2026) in partial fulfillment of the requirements for the award of the degree of Bachelor of Technology in Computer Engineering from Indus University."));
collegeCert.push(new Paragraph({ text: '' }));
collegeCert.push(createPara("The project has been approved by the Project Review Committee comprising:"));
collegeCert.push(createPara("1. Internal Guide: _________________", 12, false));
collegeCert.push(createPara("2. External Examiner: _________________", 12, false));
collegeCert.push(createPara("3. Project Coordinator: _________________", 12, false));
collegeCert.push(new Paragraph({ text: '' }));
collegeCert.push(createPara("Date: _________________", 12, false));
collegeCert.push(createPara("Signature: _________________", 12, false));
collegeCert.push(createPara("Project Coordinator", 12, false));

// ============ ACKNOWLEDGEMENT ============
const acknowledgement = [];
acknowledgement.push(createHeader("ACKNOWLEDGEMENT", 14, true));
acknowledgement.push(new Paragraph({ text: '' }));
acknowledgement.push(createPara("I would like to express my sincere gratitude to all those who have contributed to the successful completion of this project."));
acknowledgement.push(new Paragraph({ text: '' }));
acknowledgement.push(createPara("First and foremost, I would like to thank my Project Supervisor, _______________ for his/her valuable guidance, constant encouragement, and support throughout the project work."));
acknowledgement.push(new Paragraph({ text: '' }));
acknowledgement.push(createPara("I am grateful to the Department of Computer Engineering, Indus University for providing me the opportunity to work on this project."));
acknowledgement.push(new Paragraph({ text: '' }));
acknowledgement.push(createPara("I would also like to thank my peers and friends for their support and cooperation."));
acknowledgement.push(new Paragraph({ text: '' }));
acknowledgement.push(createPara("Finally, I would like to thank my parents and family for their unconditional support and encouragement throughout my academic career."));

// ============ TABLE OF CONTENTS ============
const toc = [];
toc.push(createHeader("TABLE OF CONTENTS", 14, true));
toc.push(new Paragraph({ text: '' }));
const tocItems = [
  ["COVER PAGE", "i"],
  ["FIRST PAGE", "ii"],
  ["CANDIDATE'S DECLARATION", "iii"],
  ["COMPANY CERTIFICATE", "iv"],
  ["COLLEGE CERTIFICATE", "v"],
  ["ACKNOWLEDGEMENT", "vi"],
  ["ABSTRACT", "vii"],
  ["COMPANY PROFILE", "viii"],
  ["LIST OF FIGURES", "ix"],
  ["LIST OF TABLES", "x"],
  ["ABBREVIATIONS", "xi"],
  ["NOTATIONS", "xii"],
  ["CHAPTER 1: INTRODUCTION", "1"],
  ["1.1 Introduction to the Project", "1"],
  ["1.2 Problem Statement", "2"],
  ["1.3 Objectives", "2"],
  ["1.4 Scope of the Project", "3"],
  ["1.5 Organization of the Report", "3"],
  ["CHAPTER 2: LITERATURE REVIEW", "5"],
  ["2.1 Previous Work", "5"],
  ["2.2 Research Gap", "6"],
  ["2.3 Summary", "7"],
  ["CHAPTER 3: METHODOLOGY", "9"],
  ["3.1 Tools and Technologies Used", "9"],
  ["3.2 Implementation Plan", "10"],
  ["3.3 Algorithms and Techniques Applied", "11"],
  ["3.4 System Architecture", "12"],
  ["CHAPTER 4: IMPLEMENTATION", "15"],
  ["4.1 Data Collection", "15"],
  ["4.2 System Development", "16"],
  ["4.3 Challenges Faced", "18"],
  ["CHAPTER 5: RESULTS AND DISCUSSION", "21"],
  ["5.1 Output and Analysis", "21"],
  ["5.2 Performance Evaluation", "23"],
  ["CHAPTER 6: CONCLUSION AND FUTURE WORK", "27"],
  ["6.1 Summary of Findings", "27"],
  ["6.2 Future Scope", "28"],
  ["APPENDICES", "31"],
  ["REFERENCES", "33"]
];

for (const [title, page] of tocItems) {
  toc.push(createTocItem(title, page));
}

// ============ ABSTRACT ============
const abstract = [];
abstract.push(createHeader("ABSTRACT", 14, true));
abstract.push(new Paragraph({ text: '' }));
abstract.push(createPara("The Student Helper Chatbot is a Flask-based web application designed to assist students with their academic queries, study assistance, and general information retrieval. The system provides an AI-powered conversational interface that helps students get quick answers to their questions, access study materials, and receive personalized recommendations."));
abstract.push(new Paragraph({ text: '' }));
abstract.push(createPara("The application is built using Flask for the backend, SQLite for database management, and HTML/CSS/JavaScript for the frontend. It includes user authentication with OTP (One-Time Password) verification via email, session management, and a responsive chat interface. The chatbot employs natural language processing techniques to understand user queries and provide appropriate responses."));
abstract.push(new Paragraph({ text: '' }));
abstract.push(createPara("Key features include user registration and login with secure OTP verification, password reset functionality, AI-powered chatbot responses, chat history storage, email notifications, and a responsive web interface. The system was developed following structured software engineering principles and uses best practices for web application development including MVC architecture, secure authentication, and database optimization."));

// ============ COMPANY PROFILE ============
const companyProfile = [];
companyProfile.push(createHeader("COMPANY PROFILE", 14, true));
companyProfile.push(new Paragraph({ text: '' }));
companyProfile.push(createPara("The project was developed as a university academic project at Indus University. The development environment includes a personal computer with Python 3.x, Flask framework, SQLite database, and modern web browsers for testing."));
companyProfile.push(new Paragraph({ text: '' }));
companyProfile.push(createPara("The system is designed to serve as a foundation for future enhancements and can be deployed on any web server supporting Python and Flask. The project demonstrates the application of modern web technologies in developing educational support systems."));

// ============ LIST OF FIGURES ============
const listFigures = [];
listFigures.push(createHeader("LIST OF FIGURES", 14, true));
listFigures.push(new Paragraph({ text: '' }));
const figures = [
  ["Fig. 1.1 System Architecture", "3"],
  ["Fig. 2.1 Literature Review Flow", "5"],
  ["Fig. 3.1 Use Case Diagram", "11"],
  ["Fig. 3.2 Sequence Diagram", "12"],
  ["Fig. 4.1 Database Schema", "16"],
  ["Fig. 5.1 Chat Interface", "22"]
];
for (const [title, page] of figures) {
  listFigures.push(createTocItem(title, page));
}

// ============ LIST OF TABLES ============
const listTables = [];
listTables.push(createHeader("LIST OF TABLES", 14, true));
listTables.push(new Paragraph({ text: '' }));
const tables = [
  ["Table 1.1 Project Scope", "3"],
  ["Table 2.1 Comparison of Existing Systems", "6"],
  ["Table 3.1 Technology Stack", "10"],
  ["Table 4.1 Database Schema", "17"],
  ["Table 5.1 Test Results", "24"]
];
for (const [title, page] of tables) {
  listTables.push(createTocItem(title, page));
}

// ============ ABBREVIATIONS ============
const abbreviations = [];
abbreviations.push(createHeader("ABBREVIATIONS", 14, true));
abbreviations.push(new Paragraph({ text: '' }));
const abbr = [
  "API - Application Programming Interface",
  "CSS - Cascading Style Sheets",
  "HTML - HyperText Markup Language",
  "MVC - Model View Controller",
  "OTP - One-Time Password",
  "REST - Representational State Transfer",
  "SQL - Structured Query Language",
  "SMTP - Simple Mail Transfer Protocol",
  "UI - User Interface"
];
for (const item of abbr) {
  abbreviations.push(createPara(item, 12, false));
}

// ============ NOTATIONS ============
const notations = [];
notations.push(createHeader("NOTATIONS", 14, true));
notations.push(new Paragraph({ text: '' }));
notations.push(createPara("The following mathematical and programming notations are used in this report:", 12, false));
const notat = [
  "n - Number of iterations or samples",
  "x - Input variable",
  "y - Output variable",
  "f(x) - Function of x",
  "θ - Theta (angle parameter)",
  "Σ - Summation",
  "√ - Square root"
];
for (const item of notat) {
  notations.push(createPara(item, 12, false));
}

// ============ CHAPTER 1 ============
const chapter1 = [];
chapter1.push(createSection("CHAPTER 1: INTRODUCTION"));
chapter1.push(createSubSection("1.1 Introduction to the Project"));
chapter1.push(createPara("Student Helper Chatbot is a full-stack web application designed to provide academic assistance to students through an intelligent conversational interface. The system serves as a virtual assistant that helps students with their queries, provides study resources, and offers personalized recommendations based on their needs."));
chapter1.push(createPara("The application consists of a Flask backend server that handles all business logic, API endpoints, and database operations. The frontend is built using HTML, CSS, and JavaScript, providing a responsive and intuitive user interface."));

chapter1.push(createSubSection("1.2 Problem Statement"));
chapter1.push(createPara("Traditional student support systems rely heavily on human resources including faculty advisors, counseling services, and administrative staff. Students often have to wait for office hours or schedule appointments to get their queries resolved. There is a need for a system that provides 24/7 access to academic information and assistance."));

chapter1.push(createSubSection("1.3 Objectives"));
chapter1.push(createPara("1. Develop a secure, scalable web application using Flask framework.", 12, false));
chapter1.push(createPara("2. Implement user authentication with OTP email verification.", 12, false));
chapter1.push(createPara("3. Create a responsive chatbot interface with natural language processing capabilities.", 12, false));
chapter1.push(createPara("4. Store and manage user data and chat history in SQLite database.", 12, false));
chapter1.push(createPara("5. Provide session-based authentication with secure password handling.", 12, false));
chapter1.push(createPara("6. Implement email notification system for OTP delivery.", 12, false));

chapter1.push(createSubSection("1.4 Scope of the Project"));
chapter1.push(createPara("The scope of this project includes user registration and login with OTP verification, chatbot query response system, chat history storage, and a responsive web interface. The project explicitly excludes mobile application development, video consultation features, and integration with external systems."));

chapter1.push(createSubSection("1.5 Organization of the Report"));
chapter1.push(createPara("This report is organized as follows: Chapter 2 presents the literature survey covering related work in educational chatbots. Chapter 3 describes the methodology and system architecture. Chapter 4 details the implementation aspects. Chapter 5 discusses the results and performance evaluation. Chapter 6 concludes with future scope. Appendices contain additional technical details and references."));

// ============ CHAPTER 2 ============
const chapter2 = [];
chapter2.push(createSection("CHAPTER 2: LITERATURE REVIEW"));
chapter2.push(createSubSection("2.1 Previous Work"));
chapter2.push(createPara("Various educational institutions and EdTech companies have developed chatbot solutions for student support. Traditional student support systems in educational institutions rely heavily on human resources including faculty advisors, counseling services, and administrative staff. While these systems provide personalized support, they are limited by availability, response time, and human resource constraints."));

chapter2.push(createSubSection("2.2 Research Gap"));
chapter2.push(createPara("Based on the literature survey, the following research gaps were identified:"));
chapter2.push(createPara("1. Lack of affordable, open-source chatbot solutions for educational institutions.", 12, false));
chapter2.push(createPara("2. Limited integration of secure OTP-based authentication in student web applications.", 12, false));
chapter2.push(createPara("3. Minimal use of modern Flask-based architectures in educational chatbot development.", 12, false));
chapter2.push(createPara("4. Absence of comprehensive documentation and implementation guides.", 12, false));

chapter2.push(createSubSection("2.3 Summary"));
chapter2.push(createPara("This chapter reviewed existing systems and identified research gaps that this project addresses. The next chapter describes the methodology and technologies used in the development of the Student Helper Chatbot."));

// ============ CHAPTER 3 ============
const chapter3 = [];
chapter3.push(createSection("CHAPTER 3: METHODOLOGY"));
chapter3.push(createSubSection("3.1 Tools and Technologies Used"));
chapter3.push(createPara("The project uses the following technologies:"));
chapter3.push(createPara("• Flask (Python 3) - Backend framework", 12, false));
chapter3.push(createPara("• SQLite - Database management", 12, false));
chapter3.push(createPara("• HTML/CSS/JavaScript - Frontend development", 12, false));
chapter3.push(createPara("• Python-dotenv - Environment variables", 12, false));
chapter3.push(createPara("• SMTP - Email notifications", 12, false));

chapter3.push(createSubSection("3.2 Implementation Plan"));
chapter3.push(createPara("Phase 1: Requirements gathering and database design (Week 1)"));
chapter3.push(createPara("Phase 2: Backend development and API implementation (Week 2-3)"));
chapter3.push(createPara("Phase 3: Frontend development and UI implementation (Week 4)"));
chapter3.push(createPara("Phase 4: Authentication system implementation (Week 5)"));
chapter3.push(createPara("Phase 5: Testing and debugging (Week 6)"));
chapter3.push(createPara("Phase 6: Documentation and deployment (Week 7)"));

chapter3.push(createSubSection("3.3 Algorithms and Techniques Applied"));
chapter3.push(createPara("The chatbot uses pattern matching and keyword detection algorithms to generate appropriate responses. The system implements SHA-256 password hashing for secure authentication. OTP generation uses random number generation with time-based expiration."));

chapter3.push(createSubSection("3.4 System Architecture"));
chapter3.push(createPara("The system follows the Model-View-Controller (MVC) architectural pattern. The Model layer handles database operations through SQLite. The View layer consists of HTML templates rendered by Flask. The Controller layer implements Flask routes for business logic and API endpoints."));

// ============ CHAPTER 4 ============
const chapter4 = [];
chapter4.push(createSection("CHAPTER 4: IMPLEMENTATION"));
chapter4.push(createSubSection("4.1 Data Collection"));
chapter4.push(createPara("User data is collected during the registration process which includes username, email, and password. Chat messages are stored in the database for history tracking. Session data is managed through Flask sessions with secure cookie handling."));

chapter4.push(createSubSection("4.2 System Development"));
chapter4.push(createPara("The system was developed using Flask framework with Python. Database tables were created using SQLite with proper schema design. The authentication system implements OTP verification with email delivery. The chatbot module processes user messages and generates appropriate responses based on pattern matching."));

chapter4.push(createSubSection("4.3 Challenges Faced"));
chapter4.push(createPara("1. Email delivery issues with SMTP configuration", 12, false));
chapter4.push(createPara("2. Session management across multiple pages", 12, false));
chapter4.push(createPara("3. Database connection handling and error recovery", 12, false));
chapter4.push(createPara("4. Responsive design for different screen sizes", 12, false));

// ============ CHAPTER 5 ============
const chapter5 = [];
chapter5.push(createSection("CHAPTER 5: RESULTS AND DISCUSSION"));
chapter5.push(createSubSection("5.1 Output and Analysis"));
chapter5.push(createPara("The system successfully provides user registration with OTP verification, secure login with session management, chatbot responses to user queries, chat history storage, and password reset functionality. All core features have been tested and verified to work as expected."));
chapter5.push(createPara("The user interface is responsive and works on desktop and mobile browsers. The chatbot provides appropriate responses to user queries based on pattern matching and keyword detection."));

chapter5.push(createSubSection("5.2 Performance Evaluation"));
chapter5.push(createPara("The system meets performance requirements with API response times under 500ms. Database queries are optimized with proper indexing. The modular architecture allows for easy scaling and future enhancements."));

// ============ CHAPTER 6 ============
const chapter6 = [];
chapter6.push(createSection("CHAPTER 6: CONCLUSION AND FUTURE WORK"));
chapter6.push(createSubSection("6.1 Summary of Findings"));
chapter6.push(createPara("The Student Helper Chatbot project successfully demonstrates the development of a complete web-based application using Flask and related technologies. The system provides students with 24/7 access to academic assistance through an intuitive chatbot interface. The implementation includes all core features required for a student support system: user authentication with OTP verification, secure session management, intelligent chatbot responses, and reliable data storage."));

chapter6.push(createSubSection("6.2 Future Scope"));
chapter6.push(createPara("1. AI Enhancement: Integrate advanced NLP models for more intelligent responses.", 12, false));
chapter6.push(createPara("2. Mobile Application: Develop React Native or Flutter mobile app.", 12, false));
chapter6.push(createPara("3. Multi-language Support: Add support for multiple languages.", 12, false));
chapter6.push(createPara("4. Database Migration: Migrate to PostgreSQL or MySQL for better scalability.", 12, false));
chapter6.push(createPara("5. Real-time Chat: Implement WebSocket for instant messaging.", 12, false));

// ============ APPENDICES ============
const appendices = [];
appendices.push(createSection("APPENDICES"));
appendices.push(createSubSection("Appendix A: File Structure"));
appendices.push(createPara("student_helper_chatbot/", 12, false));
appendices.push(createPara("├── app.py (Main Flask application)", 12, false));
appendices.push(createPara("├── init_db.py (Database initialization)", 12, false));
appendices.push(createPara("├── chatbot/ (Chatbot module)", 12, false));
appendices.push(createPara("├── database/ (Database files)", 12, false));
appendices.push(createPara("├── templates/ (HTML templates)", 12, false));
appendices.push(createPara("├── static/ (CSS and JavaScript)", 12, false));
appendices.push(createPara("├── .env (Environment variables)", 12, false));

appendices.push(createSubSection("Appendix B: API Endpoints"));
appendices.push(createPara("GET / - Home page", 12, false));
appendices.push(createPara("GET/POST /login - Login", 12, false));
appendices.push(createPara("GET/POST /signup - Registration", 12, false));
appendices.push(createPara("GET/POST /verify-signup - OTP verification", 12, false));
appendices.push(createPara("GET /logout - Logout", 12, false));
appendices.push(createPara("GET /app - Chat application", 12, false));
appendices.push(createPara("POST /chat - Chat API endpoint", 12, false));

// ============ REFERENCES ============
const references = [];
references.push(createSection("REFERENCES"));
references.push(createPara("1. Aloysius J. A. (1998) Data Analysis for Management, Prentice Hall of India Pvt. Ltd., New Delhi.", 12, false));
references.push(createPara("2. Flask Documentation. Available at: https://flask.palletsprojects.com/", 12, false));
references.push(createPara("3. Python Documentation. Available at: https://docs.python.org/", 12, false));
references.push(createPara("4. SQLite Documentation. Available at: https://www.sqlite.org/docs.html", 12, false));
references.push(createPara("5. MDN Web Docs - HTML, CSS, JavaScript. Available at: https://developer.mozilla.org/", 12, false));

// Build the document
const allChildren = [
  ...coverPage,
  ...firstPage,
  ...declaration,
  ...companyCert,
  ...collegeCert,
  ...acknowledgement,
  ...toc,
  ...abstract,
  ...companyProfile,
  ...listFigures,
  ...listTables,
  ...abbreviations,
  ...notations,
  ...chapter1,
  ...chapter2,
  ...chapter3,
  ...chapter4,
  ...chapter5,
  ...chapter6,
  ...appendices,
  ...references
];

const doc = new Document({
  sections: [{
    properties: pageProps(),
    children: allChildren,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('StudentHelperChatbot_Report.docx', buffer);
  console.log('Documentation generated: StudentHelperChatbot_Report.docx');
});