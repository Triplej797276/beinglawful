const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const admin = require('firebase-admin');
const session = require('express-session');
const { body, validationResult } = require('express-validator');
const nodemailer = require('nodemailer');
const schedule = require('node-schedule');
require('dotenv').config();
const multer = require('multer');
const xlsx = require('xlsx');
const crypto = require('crypto');
const fs = require('fs');
const XLSX = require('xlsx');

const saltRounds = 10;
// Multer setup for file uploads
// Multer setup for Excel file uploads
const storage = multer.memoryStorage();

const uploadpdf = multer({
  storage: storage,
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB
  fileFilter: (req, file, cb) => {
    if (file.mimetype === "application/pdf") {
      cb(null, true);
    } else {
      cb(new Error("Only PDF files are allowed"), false);
    }
  },
}).single("pdfFile");
const upload = multer({
  storage: storage,
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB
  fileFilter: (req, file, cb) => {
    const allowedTypes = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
      "application/vnd.ms-excel", // .xls
      "image/jpeg",
      "image/png",
      "image/jpg",
      "application/pdf", // ✅ added PDF
    ];

    if (allowedTypes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error("Only Excel (.xlsx, .xls), PDF, and image files are allowed"), false);
    }
  },
}).fields([
  { name: "excelFile", maxCount: 1 },
  { name: "images", maxCount: 10 },
  { name: "pdfFile", maxCount: 1 }, 
]);
const uploadImages = multer({
  limits: { fileSize: 5 * 1024 * 1024 },
  storage: multer.memoryStorage()
}).array('images', 10);
const app = express();
const uploadSchoolMedia = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB per file
  fileFilter: (req, file, cb) => {
    const allowedTypes = [
      'image/jpeg', 'image/png', 'image/jpg',
      'video/mp4', 'video/mpeg', 'video/quicktime'
    ];
    if (allowedTypes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error('Only image and video files are allowed'), false);
    }
  }
}).fields([
  { name: 'photos', maxCount: 2 },
  { name: 'videos', maxCount: 2 }
]);


//excel 

const uploadExcel = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 }, // 5MB limit
}).single('excelFile');

const uploadFeedback = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 }, // 5MB limit
}).single('feedbackExcelFile');
// Set view engine to EJS and specify the views directory
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
console.log('Views directory set to:', app.get('views'));

// Disable view cache for development
app.set('view cache', false);

// Session middleware
app.use(session({
    secret: 'your-secret-key',
    resave: false,
    saveUninitialized: false,
    cookie: { secure: false, maxAge: 24 * 60 * 60 * 1000 }
}));

// Parse JSON for fetch requests
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use('/activities', express.static('activities'));
// Trial test questions
const trialTests = {
    trial1: [
        { question: "Who is Cheif Minister of Maharashtra?", options: ["Devendra Fadnavis", "Udhav Thakarey", "Ajit Pawar", "Eknath Shinde"], correctAnswer: "Devendra Fadnavis" },
        { question: "Who is President of India?", options: ["Rahul Gandhi", "Narendra Modi", "Yogi Aditya Nath", "Dropady Murmu"], correctAnswer: "Dropady Murmu" }
    ],
    trial2: [
        { question: "What is the minimum age to vote in India?", options: ["16", "20", "18", "17"], correctAnswer: "18" },
        { question: "The Constitution of India was adopted on?", options: ["15 August 1947", "26 January 1950", "26 November 1949", "2 October 1950"], correctAnswer: "26 November 1949" }
    ]
};
    
// Validate environment variables for Firebase and Email
const requiredEnvVars = [
    'FIREBASE_PROJECT_ID',
    'FIREBASE_PRIVATE_KEY_ID',
    'FIREBASE_PRIVATE_KEY',
    'FIREBASE_CLIENT_EMAIL',
    'FIREBASE_CLIENT_ID',
    'FIREBASE_CLIENT_X509_CERT_URL',
    'EMAIL_USER',
    'EMAIL_PASS'
];
const missingEnvVars = requiredEnvVars.filter(varName => !process.env[varName]);
if (missingEnvVars.length > 0) {
    console.error(`Missing environment variables: ${missingEnvVars.join(', ')}`);
    process.exit(1);
}
let db, bucket;
try {
    const serviceAccount = {
        type: "service_account",
        project_id: process.env.FIREBASE_PROJECT_ID,
        private_key_id: process.env.FIREBASE_PRIVATE_KEY_ID,
        private_key: process.env.FIREBASE_PRIVATE_KEY.replace(/\\n/g, '\n'),
        client_email: process.env.FIREBASE_CLIENT_EMAIL,
        client_id: process.env.FIREBASE_CLIENT_ID,
        auth_uri: "https://accounts.google.com/o/oauth2/auth",
        token_uri: "https://oauth2.googleapis.com/token",
        auth_provider_x509_cert_url: "https://www.googleapis.com/oauth2/v1/certs",
        client_x509_cert_url: process.env.FIREBASE_CLIENT_X509_CERT_URL,
        universe_domain: "googleapis.com"
    };

    admin.initializeApp({
        credential: admin.credential.cert(serviceAccount),
        storageBucket: process.env.FIREBASE_STORAGE_BUCKET || 'beinglawful-ee5a4.appspot.com'
    });
    db = admin.firestore();
    bucket = admin.storage().bucket(); // Moved inside try block after initialization
    console.log('Firebase initialized successfully');
} catch (error) {
    console.error('Failed to initialize Firebase:', error.message, error.stack);
    process.exit(1);
}

// Set up Express app
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, '.')));

// Hardcoded admin credentials
const ADMIN_USERNAME = 'admin';
const ADMIN_PASSWORD = 'admin12345';

//Email setup
// const transporter = nodemailer.createTransport({
//     host: 'smtp-mail.outlook.com',
//     port: 587,
//     secure: false, // Use STARTTLS
//     auth: {
//         user: process.env.EMAIL_USER,
//         pass: process.env.EMAIL_PASS
//     },
//     tls: {
//         // Remove ciphers: 'SSLv3'
//         // Optionally specify modern TLS versions if needed
//         minVersion: 'TLSv1.2' // Ensure TLS 1.2 or higher
//     }
// });

// transporter.verify(function (error, success) {
//     if (error) {
//         console.log('SMTP Verify Error:', error);
//     } else {
//         console.log('SMTP Connection Success ✅ Ready to send emails');
//     }
// });
// Email templates
// const emailTemplates = {
//     schoolRegistration: {
//         subject: 'School Registration Confirmation - Being Lawful',
//         text: (schoolName, email, password) => `Dear ${schoolName},\n\nThank you for registering with Being Lawful! Your login credentials are:\n\nEmail: ${email}\nPassword: ${password}\n\nPlease use these to log in to your school dashboard.\n\nBest regards,\nBeing Lawful Team`,
//         html: (schoolName, email, password) => `
//             <h2>School Registration Confirmation</h2>
//             <p>Dear ${schoolName},</p>
//             <p>Thank you for registering with <strong>Being Lawful</strong>!</p>
//             <p>Your login credentials are:</p>
//             <ul>
//                 <li><strong>Email:</strong> ${email}</li>
//                 <li><strong>Password:</strong> ${password}</li>
//             </ul>
//             <p>Please use these to log in to your school dashboard.</p>
//             <p>Best regards,<br>Being Lawful Team</p>
//         `
//     },
//     trainerRegistration: {
//         subject: 'Trainer Registration Confirmation - Being Lawful',
//         text: (trainerName, email, password) => `Dear ${trainerName},\n\nThank you for registering as a trainer with Being Lawful! Your login credentials are:\n\nEmail: ${email}\nPassword: ${password}\n\nPlease use these to log in to your trainer dashboard.\n\nBest regards,\nBeing Lawful Team`,
//         html: (trainerName, email, password) => `
//             <h2>Trainer Registration Confirmation</h2>
//             <p>Dear ${trainerName},</p>
//             <p>Thank you for registering as a trainer with <strong>Being Lawful</strong>!</p>
//             <p>Your login credentials are:</p>
//             <ul>
//                 <li><strong>Email:</strong> ${email}</li>
//                 <li><strong>Password:</strong> ${password}</li>
//             </ul>
//             <p>Please use these to log in to your trainer dashboard.</p>
//             <p>Best regards,<br>Being Lawful Team</p>
//         `
//     },
//     schoolWorkshopReminder: {
//         subject: 'Workshop Reminder - Being Lawful',
//         text: (schoolName, eventDate) => `Dear ${schoolName},\n\nThis is a reminder for your upcoming workshop with Being Lawful scheduled on ${eventDate}. Please ensure all arrangements are in place.\n\nBest regards,\nBeing Lawful Team`,
//         html: (schoolName, eventDate) => `
//             <h2>Workshop Reminder</h2>
//             <p>Dear ${schoolName},</p>
//             <p>This is a reminder for your upcoming workshop with <strong>Being Lawful</strong> scheduled on <strong>${eventDate}</strong>.</p>
//             <p>Please ensure all arrangements are in place.</p>
//             <p>Best regards,<br>Being Lawful Team</p>
//         `
//     },
//     studentWorkshopReminder: {
//         subject: 'Workshop Reminder - Being Lawful',
//         text: (studentName, eventDate) => `Dear ${studentName}'s Parent,\n\nThis is a reminder for the upcoming workshop at your child's school with Being Lawful, scheduled on ${eventDate}. We look forward to your child's participation.\n\nBest regards,\nBeing Lawful Team`,
//         html: (studentName, eventDate) => `
//             <h2>Workshop Reminder</h2>
//             <p>Dear ${studentName}'s Parent,</p>
//             <p>This is a reminder for the upcoming workshop at your child's school with <strong>Being Lawful</strong>, scheduled on <strong>${eventDate}</strong>.</p>
//             <p>We look forward to your child's participation.</p>
//             <p>Best regards,<br>Being Lawful Team</p>
//         `
//     }
// };

// Send email function
// const sendEmail = async (to, subject, text, html) => {
//     try {
//         const mailOptions = {
//             from: process.env.EMAIL_USER,
//             to,
//             subject,
//             text,
//             html
//         };
//         const info = await transporter.sendMail(mailOptions);
//         console.log(`Email sent: ${info.messageId}`);
//     } catch (error) {
//         console.error(`Error sending email to ${to}:`, error.message, error.stack);
//         throw error;
//     }
// };

// Utility Functions

// Fetch random questions from Firestore
// Load questions from JSON file
// Load questions from JSON file
let MASTER_POOL;
try {
    const filePath = path.join(__dirname, 'question.json');
    const rawData = fs.readFileSync(filePath);
    MASTER_POOL = JSON.parse(rawData); // Directly parse array as per provided format
    if (!Array.isArray(MASTER_POOL) || MASTER_POOL.length < 30) {
        throw new Error('Invalid or insufficient questions in questions.json');
    }
    console.log(`Loaded ${MASTER_POOL.length} questions from questions.json`);
} catch (error) {
    console.error('Error loading questions.json:', error.message);
    process.exit(1); // Exit if questions cannot be loaded
}

// Map to store IP-based assignments: { ip: { sessionId, sets, timestamp } }
const assignments = new Map();

// Shuffle array using Fisher-Yates algorithm with a seed
function shuffleArray(array, seed) {
    if (!seed) {
        console.warn('Seed is undefined, generating fallback seed');
        seed = crypto.randomUUID();
    }
    const random = seededRandom(seed);
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
}

// Seeded random number generator for reproducible shuffling
function seededRandom(seed) {
    let x = 0, y = 0, z = 0, w = 0;

    function seedRand(s) {
        if (!s) {
            console.warn('seedRand received undefined seed, using fallback');
            s = crypto.randomUUID();
        }
        const hash = crypto.createHash('sha256').update(s).digest('hex');
        x = parseInt(hash.slice(0, 8), 16);
        y = parseInt(hash.slice(8, 16), 16);
        z = parseInt(hash.slice(16, 24), 16);
        w = parseInt(hash.slice(24, 32), 16);
    }

    seedRand(seed);

    return function () {
        const t = x ^ (x << 11);
        x = y;
        y = z;
        z = w;
        w = w ^ (w >> 19) ^ t ^ (t >> 8);
        return (w >>> 0) / 0x100000000;
    };
}

// Clean expired assignments (older than 1 hour)
function cleanExpiredAssignments() {
    const now = Date.now();
    for (const [ip, assignment] of assignments) {
        if (now - assignment.timestamp > 60 * 60 * 1000) { // 1 hour
            assignments.delete(ip);
            console.log(`Cleared expired assignment for IP: ${ip}`);
        }
    }
}

// Get random questions with validation
async function getRandomQuestions(limit = 30, seed) {
    try {
        if (!seed) {
            console.warn('getRandomQuestions received undefined seed, generating fallback');
            seed = crypto.randomUUID();
        }

        let questions = [...MASTER_POOL]; // Use MASTER_POOL

        // Filter valid questions
        questions = questions.filter(mcq => {
            const isValid =
                mcq.id && typeof mcq.id === 'string' &&
                mcq.question && typeof mcq.question === 'string' &&
                Array.isArray(mcq.options) && mcq.options.length >= 4 &&
                mcq.options.every(opt => typeof opt === 'string' && opt.trim()) &&
                mcq.correctAnswer && typeof mcq.correctAnswer === 'string' &&
                mcq.options.includes(mcq.correctAnswer);
            if (!isValid) console.warn(`Filtered out invalid MCQ: ${JSON.stringify(mcq)}`);
            return isValid;
        });

        // Check if enough valid questions are available
        if (questions.length < limit) {
            throw new Error(`Only ${questions.length} valid questions available, need ${limit}`);
        }

        // Shuffle questions using the provided seed
        questions = shuffleArray([...questions], seed).slice(0, limit);
        return questions;
    } catch (error) {
        console.error('Error fetching random questions:', error.message, error.stack);
        throw error;
    }
}

// Generate unique MCQ set
async function generateUniqueMCQSet(ip, sessionId) {
    try {
        cleanExpiredAssignments();

        // Get or initialize assignment for this IP
        let assignment = assignments.get(ip) || { sessionId, sets: [], timestamp: Date.now() };

        // Track used sets to avoid duplicates
        const usedSetIds = new Set(assignment.sets.map(set => set.seed));

        // Generate a unique seed for shuffling
        let seed;
        let attempts = 0;
        const maxAttempts = 100;
        do {
            seed = crypto.randomUUID() + ip + sessionId + Date.now();
            attempts++;
            if (attempts > maxAttempts) {
                // Fallback: Use least recently used set
                const allAssignments = Array.from(assignments.values());
                if (allAssignments.length > 0) {
                    const oldestAssignment = allAssignments.sort((a, b) => a.timestamp - b.timestamp)[0];
                    seed = oldestAssignment.sets[0].seed;
                    console.log(`Fallback: Reusing set for IP: ${ip}`);
                    break;
                } else {
                    seed = crypto.randomUUID(); // Use new seed if no assignments exist
                }
            }
        } while (usedSetIds.has(seed));

        // Get 30 random questions
        const questions = await getRandomQuestions(30, seed);

        // Generate quiz set
        const quizSet = questions.map(q => ({
            id: q.id,
            question: q.question,
            options: shuffleArray([...q.options], seed + q.id), // Shuffle options
            // correctAnswer omitted for client
        }));

        // Update assignments
        assignment.sets.push({ seed, indices: questions.map(q => MASTER_POOL.findIndex(m => m.id === q.id)) });
        assignment.sessionId = sessionId;
        assignment.timestamp = Date.now();
        assignments.set(ip, assignment);

        return { quizSet, sessionId, seed };
    } catch (error) {
        console.error('Error in generateUniqueMCQSet:', error.message);
        throw new Error('Failed to generate quiz set');
    }
}

// Express middleware to parse JSON
app.use(express.json());

// Express route to serve quiz
app.get('/start_quiz', async (req, res) => {
    const ip = req.ip; // Get client IP
    const sessionId = crypto.randomUUID(); // Generate unique session ID

    try {
        const { quizSet, sessionId: assignedSessionId } = await generateUniqueMCQSet(ip, sessionId);
        res.json({ quizSet, sessionId: assignedSessionId });
    } catch (error) {
        console.error('Error in /start_quiz:', error.message);
        res.status(500).json({ error: 'Failed to generate quiz set' });
    }
});

// Route to submit answers and validate
app.post('/submit_quiz', (req, res) => {
    const { sessionId, answers } = req.body;
    const ip = req.ip;

    try {
        const assignment = assignments.get(ip);
        if (!assignment || assignment.sessionId !== sessionId) {
            return res.status(400).json({ error: 'Invalid session' });
        }

        let score = 0;
        const results = answers.map(answer => {
            const question = MASTER_POOL.find(q => q.id === answer.questionId);
            if (!question) {
                return {
                    questionId: answer.questionId,
                    isCorrect: false,
                    correctAnswer: null
                };
            }
            const isCorrect = question.correctAnswer === answer.selectedOption;
            if (isCorrect) score++;
            return {
                questionId: answer.questionId,
                isCorrect,
                correctAnswer: question.correctAnswer
            };
        });

        res.json({ score, total: answers.length, results });
    } catch (error) {
        console.error('Error in /submit_quiz:', error.message);
        res.status(500).json({ error: 'Failed to process quiz submission' });
    }
});

// Periodic cleanup
setInterval(cleanExpiredAssignments, 10 * 60 * 1000); // Run every 10 minutes


// Fetch user by parentMobile1
async function getUserByParentMobile(parentMobile1) {
    try {
        const snapshot = await db.collection('participants')
            .where('parentMobile1', '==', parentMobile1)
            .get();
        if (snapshot.empty) {
            throw new Error(`No user found with parentMobile1: ${parentMobile1}`);
        }
        const user = snapshot.docs[0].data();
        const userId = snapshot.docs[0].id;
        return { user, userId };
    } catch (error) {
        console.error('Error in getUserByParentMobile:', error.message, error.stack);
        throw error;
    }
}

// Fetch event date details for a school
async function getEventDateDetails(schoolName) {
    try {
        console.log(`Fetching event date details for schoolName: ${schoolName}`);
        if (!schoolName) {
            console.warn('schoolName is empty or undefined');
            return { isEventDate: false, isOnOrAfterEventDate: false, eventDateMissing: true, eventDate: null };
        }

        const schoolSnapshot = await db.collection('schools')
            .where('schoolName', '==', schoolName)
            .get();
        
        if (schoolSnapshot.empty) {
            console.warn(`No school found with schoolName: ${schoolName}`);
            return { isEventDate: false, isOnOrAfterEventDate: false, eventDateMissing: true, eventDate: null };
        }

        const schoolData = schoolSnapshot.docs[0].data();
        const eventDateRaw = schoolData.eventDate;
        console.log(`Raw eventDate from Firestore: ${eventDateRaw}, Type: ${typeof eventDateRaw}`);

        if (!eventDateRaw || typeof eventDateRaw.toDate !== 'function') {
            console.warn('eventDate is null, undefined, or not a Firestore Timestamp:', eventDateRaw);
            return { isEventDate: false, isOnOrAfterEventDate: false, eventDateMissing: true, eventDate: null };
        }

        const eventDate = eventDateRaw.toDate();
        console.log(`Converted eventDate: ${eventDate}`);

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        eventDate.setHours(0, 0, 0, 0);

        console.log(`Today's date (normalized): ${today}`);
        console.log(`Event date (normalized): ${eventDate}`);

        const isEventDate = eventDate.getTime() === today.getTime();
        const isOnOrAfterEventDate = today.getTime() >= eventDate.getTime();
        console.log(`isEventDate result: ${isEventDate}`);
        console.log(`isOnOrAfterEventDate result: ${isOnOrAfterEventDate}`);

        const formattedEventDate = eventDate.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });

        return { isEventDate, isOnOrAfterEventDate, eventDateMissing: false, eventDate: formattedEventDate };
    } catch (error) {
        console.error('Error calculating event date details:', error.message, error.stack);
        return { isEventDate: false, isOnOrAfterEventDate: false, eventDateMissing: true, eventDate: null };
    }
}

// Middleware

// Check admin authentication
function requireAdmin(req, res, next) {
    if (req.session.isAdmin) {
        return next();
    }
    res.redirect('/admin-login');
}

// Check student authentication
function requireStudentAuth(req, res, next) {
    if (req.session.parentMobile1) {
        return next();
    }
    res.redirect('/login?error=Please%20login%20to%20access%20this%20page');
}

// Check if it's on or after the event date
    async function checkEventDate(req, res, next) {
        try {
            const parentMobile1 = req.session.parentMobile1 || req.params.parentMobile1;
            const { user } = await getUserByParentMobile(parentMobile1);
            const { isEventDate, isOnOrAfterEventDate, eventDateMissing, eventDate } = await getEventDateDetails(user.schoolNameDropdown || '');
            res.locals.isEventDate = isEventDate;
            res.locals.isOnOrAfterEventDate = isOnOrAfterEventDate;
            res.locals.eventDateMissing = eventDateMissing;
            res.locals.eventDate = eventDate;
            res.locals.user = user;
            next();
        } catch (error) {
            console.error('Error in checkEventDate middleware:', error.message, error.stack);
            res.locals.isEventDate = false;
            res.locals.isOnOrAfterEventDate = false;
            res.locals.eventDateMissing = true;
            res.locals.eventDate = null;
            next();
        }
    }

    // Send workshop reminder emails (runs daily)
    const checkAndSendWorkshopEmails = async () => {
        const today = new Date().toISOString().split('T')[0];
        console.log(`Checking for workshop reminders on ${today}`);

        const schoolsSnapshot = await db.collection('schools').where('isApproved', '==', true).get();
        console.log(`Found ${schoolsSnapshot.size} approved schools`);
        if (!schoolsSnapshot.empty) {
            for (const doc of schoolsSnapshot.docs) {
                const school = doc.data();
                const schoolName = school.schoolName;
                const schoolEmail = school.schoolEmail;

                if (school.eventDate && typeof school.eventDate.toDate === 'function') {
                    const eventDate = school.eventDate.toDate();
                    const eventDateString = eventDate.toISOString().split('T')[0];
                    console.log(`School ${schoolName} has event date ${eventDateString}`);

                    if (today === eventDateString) {
                        const formattedEventDate = eventDate.toLocaleDateString('en-US', {
                            year: 'numeric',
                            month: 'long',
                            day: 'numeric'
                        });
                        console.log(`Sending workshop reminder to ${schoolEmail} for event on ${formattedEventDate}`);
                        await sendEmail(
                            schoolEmail,
                            emailTemplates.schoolWorkshopReminder.subject,
                            emailTemplates.schoolWorkshopReminder.text(schoolName, formattedEventDate),
                            emailTemplates.schoolWorkshopReminder.html(schoolName, formattedEventDate)
                        );
                    }
                } else {
                    console.log(`School ${schoolName} has no event date`);
                }
            }
        } else {
            console.log('No approved schools found');
        }

        const studentsSnapshot = await db.collection('participants').where('hasCompletedMCQ', '==', true).get();
        console.log(`Found ${studentsSnapshot.size} students who completed MCQ`);
        if (!studentsSnapshot.empty) {
            for (const doc of studentsSnapshot.docs) {
                const student = doc.data();
                const studentName = student.studentName;
                const parentEmail = student.parentEmail;
                const schoolName = student.schoolNameDropdown;

                const schoolSnapshot = await db.collection('schools')
                    .where('schoolName', '==', schoolName)
                    .get();
                if (schoolSnapshot.empty) {
                    console.log(`No school found for student ${studentName} with schoolName ${schoolName}`);
                    continue;
                }

                const schoolData = schoolSnapshot.docs[0].data();
                if (schoolData.eventDate && typeof schoolData.eventDate.toDate === 'function') {
                    const eventDate = schoolData.eventDate.toDate();
                    const eventDateString = eventDate.toISOString().split('T')[0];
                    console.log(`Student ${studentName} linked to school ${schoolName} with event date ${eventDateString}`);

                    if (today === eventDateString && parentEmail) {
                        const formattedEventDate = eventDate.toLocaleDateString('en-US', {
                            year: 'numeric',
                            month: 'long',
                            day: 'numeric'
                        });
                        console.log(`Sending workshop reminder to ${parentEmail} for event on ${formattedEventDate}`);
                        // await sendEmail(
                        //     parentEmail,
                        //     emailTemplates.studentWorkshopReminder.subject,
                        //     emailTemplates.studentWorkshopReminder.text(studentName, formattedEventDate),
                        //     emailTemplates.studentWorkshopReminder.html(studentName, formattedEventDate)
                        // );
                    }
                } else {
                    console.log(`School ${schoolName} for student ${studentName} has no event date`);
                }
            }
        } else {
            console.log('No students found who completed MCQ');
        }
    };

    // Schedule daily email reminders at 10:30 PM IST
    const IST_OFFSET = 5.5 * 60 * 60 * 1000; // IST is UTC+5:30
    const job = schedule.scheduleJob('30 22 * * *', () => {
        const now = new Date();
        const istTime = new Date(now.getTime() + IST_OFFSET);
        console.log(`Running email reminder job at ${istTime.toISOString()} IST (Server time: ${now.toISOString()})`);
        checkAndSendWorkshopEmails();
    });

// Routes

    // Home route (serves static HTML)
    app.get('/', (req, res) => {
        res.sendFile(path.join(__dirname, 'index.html'));
    });

// Login page (renders login.ejs)
    app.get('/login', (req, res) => {
        const error = req.query.error || null;
        res.render('login', { error });
    });
    
app.post('/student-login', async (req, res) => {
    const { username, password } = req.body;
    try {
        if (!username || !password) {
            return res.render('login', { error: 'Username and password are required.' });
        }

        console.log(`Student login attempt for username: ${username}`);
        const { user, userId } = await getUserByParentMobile(username);

        if (!user) {
            return res.render('login', { error: 'Student not found.' });
        }

        if (!user.isApproved) {
            return res.render('login', { error: 'Approval pending. Please contact your school.' });
        }

        if (!user.birthdate || !/^\d{4}-\d{2}-\d{2}$/.test(user.birthdate)) {
            console.error(`Invalid birthdate for user ${username}: ${user.birthdate}`);
            return res.render('login', { error: 'Invalid user data. Contact administrator.' });
        }

        const [year, month, day] = user.birthdate.split('-');
        const expectedPassword = `${day}${month}${year}`;
        if (password !== expectedPassword) {
            console.log(`Invalid password for username: ${username}. Expected: ${expectedPassword}, Got: ${password}`);
            return res.render('login', { error: 'Invalid username or password.' });
        }

        req.session.parentMobile1 = username;
        console.log(`Login successful for ${username}`);

        res.redirect(`/dashboard/${encodeURIComponent(username)}`);
    } catch (error) {
        console.error('Error during student login:', error.message, error.stack);
        res.render('login', { error: 'Login failed. Please try again later.' });
    }
});


    // Dashboard (renders dashboard.ejs)
 app.get('/dashboard/:parentMobile1', requireStudentAuth, checkEventDate, async (req, res) => {
    try {
        const parentMobile1 = req.session.parentMobile1;

        const snapshot = await db.collection('participants')
            .where('parentMobile1', '==', parentMobile1)
            .get();

        if (snapshot.empty) {
            return res.redirect('/login?error=Student%20not%20found');
        }

        const doc = snapshot.docs[0];
        const user = doc.data();
        const docId = doc.id;

        let mcqs = [];
        const hasCompletedMCQ = user.hasCompletedMCQ || false;

        if (!hasCompletedMCQ) {
            await db.runTransaction(async (transaction) => {
                const userRef = db.collection('participants').doc(docId);
                const userDoc = await transaction.get(userRef);
                const userData = userDoc.data();
                mcqs = userData.currentMcqs || [];
                if (mcqs.length === 0) {
                    mcqs = await getRandomQuestions(30);
                    transaction.update(userRef, { currentMcqs: mcqs });
                }
            });
        } else {
            mcqs = user.currentMcqs || [];
        }

        const mediaSnapshot = await db.collection('participants')
            .doc(docId)
            .collection('mediaUploads')
            .orderBy('uploadedAt', 'desc')
            .get();

        const mediaUploads = mediaSnapshot.docs.map(doc => doc.data());

        console.log('Dashboard Data:', {
            studentName: user.studentName,
            parentMobile1,
            hasCompletedMCQ,
            mcqsCount: mcqs.length
        });

        res.render('dashboard', {
            studentName: user.studentName || 'N/A',
            parentMobile1,
            isChampion: user.isChampion || false,
            championMessage: user.championMessage || '',
            hasCompletedMCQ,
            hasCompletedTrial1: user.hasCompletedTrial1 || false,
            hasCompletedTrial2: user.hasCompletedTrial2 || false,
            trial1: trialTests.trial1,
            trial2: trialTests.trial2,
            mcqs,
            showResults: hasCompletedMCQ,
            score: user.score || 0,
            totalQuestions: user.totalQuestions || 30,
            percentage: user.percentage || 0,
            mediaUploads,
            showMediaSection: hasCompletedMCQ && (req.isOnOrAfterEventDate || false),
            eventDate: req.eventDate || '',
            eventDateMissing: req.eventDateMissing || true,
            isOnOrAfterEventDate: req.isOnOrAfterEventDate || false,
            isEventDate: req.isEventDate || false,
            batch: user.batch || 'N/A',
            error: null
        });
    } catch (error) {
        console.error('Error loading dashboard:', error.message, error.stack);
        res.redirect('/login?error=Dashboard%20load%20failed');
    }
});


    // MCQ test (renders mcq.ejs)
    app.get('/mcq-test/:parentMobile1', requireStudentAuth, async (req, res) => {
        try {
            const parentMobile1 = req.session.parentMobile1;
            const { user, userId } = await getUserByParentMobile(parentMobile1);

            if (user.hasCompletedMCQ) {
                return res.status(400).send('You have already completed the MCQ test.');
            }
            if (!(user.hasCompletedTrial1 || false) || !(user.hasCompletedTrial2 || false)) {
                return res.status(400).send('Please complete both trial tests before starting the main exam.');
            }

            let mcqs = [];
            await db.runTransaction(async (transaction) => {
                const userRef = db.collection('participants').doc(userId);
                const userDoc = await transaction.get(userRef);
                const userData = userDoc.data();
                mcqs = userData.currentMcqs || [];
                if (mcqs.length === 0) {
                    mcqs = await getRandomQuestions(30);
                    transaction.update(userRef, { currentMcqs: mcqs });
                }
            });

            res.render('mcq', { parentMobile1, mcqs });
        } catch (error) {
            console.error('Error in MCQ test route:', error.message, error.stack);
            res.status(500).send('Error loading MCQ test.');
        }
    });

// Submit Trial 1 (renders dashboard.ejs)
    app.post('/submit-trial1', requireStudentAuth, async (req, res) => {
        try {
            const { parentMobile1, ...answers } = req.body;
            const { user, userId } = await getUserByParentMobile(parentMobile1);

            let score = 0, correctAnswers = 0, wrongAnswers = 0;
            const trial1 = trialTests.trial1;
            trial1.forEach((mcq, index) => {
                const userAnswer = answers[`q${index}`]?.trim().toLowerCase();
                const correctAnswer = mcq.correctAnswer?.trim().toLowerCase();
                const isCorrect = userAnswer === correctAnswer;
                if (isCorrect) {
                    score++;
                    correctAnswers++;
                } else {
                    wrongAnswers++;
                }
            });

            const totalQuestions = trial1.length;
            const percentage = Math.round((score / totalQuestions) * 100);
            const mcqs = await getRandomQuestions(30);

            await db.collection('participants').doc(userId).update({
                hasCompletedTrial1: true,
                trial1Score: score,
                trial1TotalQuestions: totalQuestions,
                trial1Percentage: percentage,
                trial1CorrectAnswers: correctAnswers,
                trial1WrongAnswers: wrongAnswers,
                currentMcqs: mcqs
            });

            const updatedUser = (await db.collection('participants').doc(userId).get()).data();
            const { isEventDate, isOnOrAfterEventDate, eventDateMissing, eventDate } = await getEventDateDetails(updatedUser.schoolNameDropdown || '');

            res.render('dashboard', {
                studentName: updatedUser.studentName || 'Unknown Student',
                parentMobile1,
                hasCompletedMCQ: updatedUser.hasCompletedMCQ || false,
                hasCompletedTrial1: true,
                hasCompletedTrial2: updatedUser.hasCompletedTrial2 || false,
                mcqs,
                trial1: trialTests.trial1,
                trial2: trialTests.trial2,
                isEventDate,
                isOnOrAfterEventDate,
                eventDateMissing,
                eventDate,
                showResults: updatedUser.hasCompletedMCQ || false,
                score: updatedUser.score || 0,
                totalQuestions: updatedUser.totalQuestions || 30,
                percentage: updatedUser.percentage || 0
            });
        } catch (error) {
            console.error('Error in submit-trial1 route:', error.message, error.stack);
            res.status(500).send('Error processing trial test.');
        }
    });

    // Submit Trial 2 (renders dashboard.ejs)
    app.post('/submit-trial2', requireStudentAuth, async (req, res) => {
        try {
            const { parentMobile1, ...answers } = req.body;
            const { user, userId } = await getUserByParentMobile(parentMobile1);

            let score = 0, correctAnswers = 0, wrongAnswers = 0;
            const trial2 = trialTests.trial2;
            trial2.forEach((mcq, index) => {
                const userAnswer = answers[`q${index}`]?.trim().toLowerCase();
                const correctAnswer = mcq.correctAnswer?.trim().toLowerCase();
                const isCorrect = userAnswer === correctAnswer;
                if (isCorrect) {
                    score++;
                    correctAnswers++;
                } else {
                    wrongAnswers++;
                }
            });

            const totalQuestions = trial2.length;
            const percentage = Math.round((score / totalQuestions) * 100);
            const mcqs = await getRandomQuestions(30);

            await db.collection('participants').doc(userId).update({
                hasCompletedTrial2: true,
                trial2Score: score,
                trial2TotalQuestions: totalQuestions,
                trial2Percentage: percentage,
                trial2CorrectAnswers: correctAnswers,
                trial2WrongAnswers: wrongAnswers,
                currentMcqs: mcqs
            });

            const updatedUser = (await db.collection('participants').doc(userId).get()).data();
            const { isEventDate, isOnOrAfterEventDate, eventDateMissing, eventDate } = await getEventDateDetails(updatedUser.schoolNameDropdown || '');

            res.render('dashboard', {
                studentName: updatedUser.studentName || 'Unknown Student',
                parentMobile1,
                hasCompletedMCQ: updatedUser.hasCompletedMCQ || false,
                hasCompletedTrial1: updatedUser.hasCompletedTrial1 || false,
                hasCompletedTrial2: true,
                mcqs,
                trial1: trialTests.trial1,
                trial2: trialTests.trial2,
                isEventDate,
                isOnOrAfterEventDate,
                eventDateMissing,
                eventDate,
                showResults: updatedUser.hasCompletedMCQ || false,
                score: updatedUser.score || 0,
                totalQuestions: updatedUser.totalQuestions || 30,
                percentage: updatedUser.percentage || 0
            });
        } catch (error) {
            console.error('Error in submit-trial2 route:', error.message, error.stack);
            res.status(500).send('Error processing trial test.');
        }
    });

    // Submit MCQ (renders dashboard.ejs)
    app.post('/submit-mcq', requireStudentAuth, async (req, res) => {
        try {
            const { parentMobile1, ...answers } = req.body;
            const { user, userId } = await getUserByParentMobile(parentMobile1);
            let mcqs = user.currentMcqs || [];

            if (mcqs.length === 0) {
                return res.status(400).send('No questions found for this test session.');
            }

            let score = 0, correctAnswers = 0, wrongAnswers = 0;
            mcqs.forEach((mcq, index) => {
                const userAnswer = answers[`q${index}`]?.trim().toLowerCase();
                const correctAnswer = mcq.correctAnswer?.trim().toLowerCase();
                const isCorrect = userAnswer === correctAnswer;
                if (isCorrect) {
                    score++;
                    correctAnswers++;
                } else {
                    wrongAnswers++;
                }
            });

            const totalQuestions = mcqs.length;
            const percentage = Math.round((score / totalQuestions) * 100);
            const newMcqs = await getRandomQuestions(30);

            await db.collection('participants').doc(userId).update({
                hasCompletedMCQ: true,
                score,
                totalQuestions,
                correctAnswers,
                wrongAnswers,
                percentage,
                completedAt: admin.firestore.FieldValue.serverTimestamp(),
                currentMcqs: newMcqs
            });

            const updatedUser = (await db.collection('participants').doc(userId).get()).data();
            const { isEventDate, isOnOrAfterEventDate, eventDateMissing, eventDate } = await getEventDateDetails(updatedUser.schoolNameDropdown || '');

            res.render('dashboard', {
                studentName: updatedUser.studentName || 'Unknown Student',
                parentMobile1,
                hasCompletedMCQ: true,
                hasCompletedTrial1: updatedUser.hasCompletedTrial1 || false,
                hasCompletedTrial2: updatedUser.hasCompletedTrial2 || false,
                mcqs: newMcqs,
                trial1: trialTests.trial1,
                trial2: trialTests.trial2,
                isEventDate,
                isOnOrAfterEventDate,
                eventDateMissing,
                eventDate,
                showResults: true,
                score,
                totalQuestions,
                percentage
            });
        } catch (error) {
            console.error('Error in submit-mcq route:', error.message, error.stack);
            res.status(500).send('Error processing your exam submission.');
        }
    });

// Game Zone (renders gamezone.ejs)
app.get('/gamezone/:parentMobile1', requireStudentAuth, checkEventDate, async (req, res) => {
    try {
        const parentMobile1 = req.session.parentMobile1 || req.params.parentMobile1;

        const snapshot = await db.collection('participants')
            .where('parentMobile1', '==', parentMobile1)
            .limit(1)
            .get();

        if (snapshot.empty) {
            return res.status(404).send('User not found.');
        }

        const user = snapshot.docs[0].data();

        if (!user.hasCompletedMCQ) {
            return res.status(400).send('Please complete the main exam to access the Game Zone.');
        }

        if (res.locals.eventDateMissing) {
            return res.status(400).send('Event date yet to decide.');
        }

        if (!res.locals.isOnOrAfterEventDate) {
            return res.status(400).send(`Game Zone is only available on or after the event date: ${res.locals.eventDate}.`);
        }

        res.render('gamezone', {
            studentName: user.studentName || 'Unknown Student',
            parentMobile1
        });

    } catch (error) {
        console.error('Error in gamezone route:', error.message, error.stack);
        res.status(500).send('Error loading Game Zone.');
    }
});
    //student dashboard upload media
 app.post('/student-dashboard/upload-media', uploadSchoolMedia, async (req, res) => {
  try {
    const { parentMobile1, mediaDescription, mediaLink } = req.body;
    const photos = req.files['photos'] || [];
    const videos = req.files['videos'] || [];

    if (!parentMobile1) {
      return res.status(400).send('Parent Mobile Number is required.');
    }

    const participantSnapshot = await db.collection('participants')
      .where('parentMobile1', '==', parentMobile1)
      .get();

    if (participantSnapshot.empty) {
      return res.status(404).send('Student not found.');
    }

    const participantId = participantSnapshot.docs[0].id;
    const participantData = participantSnapshot.docs[0].data();

    if (mediaDescription && mediaDescription.length > 1000) {
      return res.status(400).send('Media description cannot exceed 1000 words.');
    }

    const bucket = admin.storage().bucket(); // Use default bucket from initialization
    const mediaUploads = [];

    for (const file of photos) {
      const fileName = `studentMedia/${participantId}/photos/${Date.now()}_${file.originalname}`;
      const fileUpload = bucket.file(fileName);

      await fileUpload.save(file.buffer, { metadata: { contentType: file.mimetype } });
      const [url] = await fileUpload.getSignedUrl({ action: 'read', expires: '03-09-2491' });

      mediaUploads.push({
        url,
        type: 'image',
        path: fileName,
        description: mediaDescription || '',
        link: mediaLink || '',
        uploadedAt: admin.firestore.FieldValue.serverTimestamp()
      });
    }

    for (const file of videos) {
      const fileName = `studentMedia/${participantId}/videos/${Date.now()}_${file.originalname}`;
      const fileUpload = bucket.file(fileName);

      await fileUpload.save(file.buffer, { metadata: { contentType: file.mimetype } });
      const [url] = await fileUpload.getSignedUrl({ action: 'read', expires: '03-09-2491' });

      mediaUploads.push({
        url,
        type: 'video',
        path: fileName,
        description: mediaDescription || '',
        link: mediaLink || '',
        uploadedAt: admin.firestore.FieldValue.serverTimestamp()
      });
    }

    if (mediaUploads.length > 0) {
      const batch = db.batch();
      mediaUploads.forEach(upload => {
        const mediaRef = db.collection('participants').doc(participantId).collection('mediaUploads').doc();
        batch.set(mediaRef, upload);
      });
      await batch.commit();
    }

    const uploadedMediaSnapshot = await db.collection('participants').doc(participantId).collection('mediaUploads').orderBy('uploadedAt', 'desc').get();
    const allUploadedMedia = uploadedMediaSnapshot.docs.map(doc => doc.data());

    const updatedUser = (await db.collection('participants').doc(participantId).get()).data();
    const { isEventDate, isOnOrAfterEventDate, eventDateMissing, eventDate } = await getEventDateDetails(updatedUser.schoolNameDropdown || '');

    res.render('dashboard', {
      studentName: updatedUser.studentName || 'Unknown Student',
      parentMobile1,
      hasCompletedMCQ: updatedUser.hasCompletedMCQ || false,
      hasCompletedTrial1: updatedUser.hasCompletedTrial1 || false,
      hasCompletedTrial2: updatedUser.hasCompletedTrial2 || false,
      mcqs: updatedUser.currentMcqs || [],
      trial1: trialTests.trial1,
      trial2: trialTests.trial2,
      isEventDate,
      isOnOrAfterEventDate,
      eventDateMissing,
      eventDate,
      showResults: updatedUser.hasCompletedMCQ || false,
      score: updatedUser.score || 0,
      totalQuestions: updatedUser.totalQuestions || 30,
      percentage: updatedUser.percentage || 0,
      mediaUploads: allUploadedMedia,
      error: null
    });
  } catch (error) {
    console.error('Error uploading media:', error);
    res.status(500).send(`Error uploading media: ${error.message}`);
  }
});

app.get('/student-dashboard/media-upload/:parentMobile1', requireStudentAuth, async (req, res) => {
    try {
        const { parentMobile1 } = req.params;
        const user = res.locals.user;

        const participantSnapshot = await db.collection('participants')
            .where('parentMobile1', '==', parentMobile1)
            .get();

        if (participantSnapshot.empty) {
            return res.status(404).send('Student not found.');
        }

        const participantDocId = participantSnapshot.docs[0].id;
        const participantData = participantSnapshot.docs[0].data();

        const mediaSnapshot = await db.collection('participants')
            .doc(participantDocId)
            .collection('mediaUploads')
            .orderBy('uploadedAt', 'desc')
            .get();

        const participantMediaUploads = mediaSnapshot.docs.map(doc => doc.data());

        const { isEventDate, isOnOrAfterEventDate, eventDateMissing, eventDate } = await getEventDateDetails(participantData.schoolNameDropdown || '');

        res.render('dashboard', {
            studentName: participantData.studentName || 'Unknown Student',
            parentMobile1,
            hasCompletedMCQ: participantData.hasCompletedMCQ || false,
            hasCompletedTrial1: participantData.hasCompletedTrial1 || false,
            hasCompletedTrial2: participantData.hasCompletedTrial2 || false,
            mcqs: participantData.currentMcqs || [],
            trial1: trialTests.trial1,
            trial2: trialTests.trial2,
            isEventDate,
            isOnOrAfterEventDate,
            eventDateMissing,
            eventDate,
            showResults: participantData.hasCompletedMCQ || false,
            score: participantData.score || 0,
            totalQuestions: participantData.totalQuestions || 30,
            percentage: participantData.percentage || 0,
            mediaUploads: participantMediaUploads,
            showMediaSection: true
        });

    } catch (error) {
        console.error('Error loading media upload page:', error);
        res.redirect('/dashboard/' + req.params.parentMobile1);
    }
});


// Certificate (renders certificate.ejs)
    app.get('/certificate/:parentMobile1', requireStudentAuth, async (req, res) => {
        try {
            const parentMobile1 = req.session.parentMobile1;
            const { user } = await getUserByParentMobile(parentMobile1);

            if (!(user.hasCompletedMCQ || false)) {
                return res.status(400).send('Please complete the main exam to access the certificate.');
            }

            const schoolSnapshot = await db.collection('schools')
                .where('schoolName', '==', user.schoolNameDropdown || '')
                .get();
            const schoolName = schoolSnapshot.empty ? 'N/A' : schoolSnapshot.docs[0].data().schoolName;

            res.render('certificate', {
                studentName: user.studentName || 'Unknown Student',
                schoolName,
                percentage: user.percentage || 0,
                completedAt: user.completedAt ? user.completedAt.toDate().toLocaleDateString() : 'N/A',
                parentMobile1
            });
        } catch (error) {
            console.error('Error in certificate route:', error.message, error.stack);
            res.status(500).send('Error loading certificate.');
        }
    });

// ... (Existing imports and setup remain unchanged)

// Book Order Route
    app.post('/submit-book-order',
        [
            body('buyerName')
                .trim()
                .notEmpty()
                .withMessage('Buyer name is required')
                .matches(/^[A-Za-z\s]{2,}$/)
                .withMessage('Invalid name format'),
            body('email')
                .trim()
                .isEmail()
                .withMessage('Invalid email address'),
            body('phone')
                .trim()
                .matches(/^\d{10}$/)
                .withMessage('Phone number must be 10 digits'),
            body('address')
                .trim()
                .isLength({ min: 5 })
                .withMessage('Address must be at least 5 characters'),
            body('city')
                .trim()
                .matches(/^[A-Za-z\s]{2,}$/)
                .withMessage('Invalid city name'),
            body('pincode')
                .trim()
                .matches(/^\d{6}$/)
                .withMessage('Pincode must be 6 digits'),
            body('quantity')
                .isInt({ min: 1, max: 10 })
                .withMessage('Quantity must be between 1 and 10'),
        ],
        async (req, res) => {
            try {
                // Check for validation errors
                const errors = validationResult(req);
                if (!errors.isEmpty()) {
                    console.error('Validation errors:', errors.array());
                    return res.status(400).render('orderError', {
                        error: errors.array()[0].msg,
                        values: req.body,
                    });
                }

                const { buyerName, email, phone, address, city, pincode, quantity } = req.body;

                // Check for duplicate order by email or phone
                const ordersRef = db.collection('orders');
                const emailSnapshot = await ordersRef.where('email', '==', email).get();
                const phoneSnapshot = await ordersRef.where('phone', '==', phone).get();

                if (!emailSnapshot.empty) {
                    return res.status(400).render('orderError', {
                        error: 'An order with this email already exists.',
                        values: req.body,
                    });
                }
                if (!phoneSnapshot.empty) {
                    return res.status(400).render('orderError', {
                        error: 'An order with this phone number already exists.',
                        values: req.body,
                    });
                }

                // Save order to Firestore
                const orderData = {
                    buyerName,
                    email,
                    phone,
                    address,
                    city,
                    pincode,
                    quantity: parseInt(quantity),
                    createdAt: admin.firestore.FieldValue.serverTimestamp(),
                };

                const orderRef = await ordersRef.add(orderData);
                console.log(`Order created with ID: ${orderRef.id}`);

                // Optionally send confirmation email (using existing nodemailer setup)
                // await sendEmail(
                //     email,
                //     'Book Order Confirmation - Being Lawful',
                //     `Dear ${buyerName},\n\nThank you for your book order! Your order ID is ${orderRef.id}.\n\nBest regards,\nBeing Lawful Team`,
                //     `
                //         <h2>Book Order Confirmation</h2>
                //         <p>Dear ${buyerName},</p>
                //         <p>Thank you for your book order with <strong>Being Lawful</strong>!</p>
                //         <p>Your order ID is: <strong>${orderRef.id}</strong></p>
                //         <p>Best regards,<br>Being Lawful Team</p>
                //     `
                // );

                // Render confirmation page
                res.render('orderConfirmation', {
                    buyerName,
                    email,
                    orderId: orderRef.id,
                });
            } catch (error) {
                console.error('Error in submit-book-order route:', error.message, error.stack);
                res.status(500).render('orderError', {
                    error: 'Failed to submit order. Please try again later.',
                    values: req.body,
                });
            }
        }
    );
    app.get('/buy-book', (req, res) => {
        try {
            res.render('buy-Book', { error: null }); // Render buy-book.ejs from views
        } catch (error) {
            console.error('Error rendering buy-book:', error.message, error.stack);
            res.status(500).render('orderError', { error: 'Failed to load the order form. Please try again later.' });
        }
    });

app.post('/admin/assign-event-date-school/:id', requireAdmin, async (req, res) => {
  try {
    const schoolId = req.params.id;
    // Use final date sent from frontend
    const eventDate = req.body.eventDateFinal || req.body.eventDate;

    if (!eventDate) {
      return res.status(400).send('Event date is required.');
    }

    const parsedDate = new Date(eventDate);
    if (isNaN(parsedDate.getTime())) {
      return res.status(400).send('Invalid event date format.');
    }

    // Save date as Firestore Timestamp
    await db.collection('schools').doc(schoolId).update({
      eventDate: admin.firestore.Timestamp.fromDate(parsedDate)
    });

    res.redirect('/admin-dashboard');
  } catch (error) {
    console.error('Error in assign-event-date-school route:', error.message, error.stack);
    res.status(500).send('Error assigning event date.');
  }
});

// Upload multiple images one by one to the specified path
    app.post('/upload-image', requireAdmin, uploadImages, async (req, res) => {
    try {
        console.log('🟢 Upload started', req.files);

        const files = req.files;
        if (!files || files.length === 0) {
        console.log('❌ No image files provided');
        return res.status(400).json({ error: 'No image files provided' });
        }

        const uploadedFiles = [];
        for (const file of files) {
        const sanitizedFileName = file.originalname.replace(/[^a-zA-Z0-9.\-_]/g, '');
        const timestamp = Date.now();
        const storagePath = `media/admin upload photo/${timestamp}_${sanitizedFileName}`;
        const storageFile = bucket.file(storagePath);

        await new Promise((resolve, reject) => {
            const stream = storageFile.createWriteStream({
            metadata: { contentType: file.mimetype },
            });

            stream.on('error', (error) => {
            console.error(`❌ Upload error for ${sanitizedFileName}:`, error);
            reject(error);
            });

            stream.on('finish', async () => {
            const [url] = await storageFile.getSignedUrl({
                action: 'read',
                expires: new Date(Date.now() + 365 * 24 * 60 * 60 * 1000),
            });

            const mediaRef = await db.collection('admin upload photo').add({
                storagePath,
                url,
                uploadedAt: admin.firestore.FieldValue.serverTimestamp(),
            });
            uploadedFiles.push({ storagePath, downloadUrl: url, mediaId: mediaRef.id });
            resolve();
            });

            stream.end(file.buffer);
        });
        }

        res.status(200).json({
        message: 'Images uploaded successfully',
        files: uploadedFiles,
        });
    } catch (error) {
        console.error('❌ Server error:', error);
        res.status(500).json({ error: 'Server error' });
    }
    });

    // Backend: /post-to-website
app.post('/post-to-website', requireAdmin, uploadImages, async (req, res) => {
  try {
    const files = req.files;
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No media files provided' });
    }

    // Validate file types and sizes
    const allowedTypes = ['image/jpeg', 'image/png', 'video/mp4'];
    const maxFileSize = 10 * 1024 * 1024; // 10MB
    for (const file of files) {
      if (!allowedTypes.includes(file.mimetype)) {
        return res.status(400).json({ error: `Invalid file type: ${file.mimetype}` });
      }
      if (file.size > maxFileSize) {
        return res.status(400).json({ error: `File too large: ${file.originalname}` });
      }
    }

    const uploadedFiles = [];
    for (const file of files) {
      const sanitizedFileName = file.originalname.replace(/[^a-zA-Z0-9.\-_]/g, '');
      const timestamp = Date.now();
      const storagePath = `gallery/${timestamp}_${sanitizedFileName}`;
      const storageFile = bucket.file(storagePath);

      await new Promise((resolve, reject) => {
        const stream = storageFile.createWriteStream({
          metadata: { contentType: file.mimetype },
        });

        stream.on('error', (err) => {
          console.error('❌ Stream error:', err.message);
          reject(err);
        });

        stream.on('finish', async () => {
          try {
            const [url] = await storageFile.getSignedUrl({
              action: 'read',
              expires: new Date(Date.now() + 365 * 24 * 60 * 60 * 1000), // 1-year expiration
            });

            // Store in 'gallery' collection with validated fields
            const mediaData = {
              storagePath: storagePath, // Path in storage bucket
              url: url, // Signed URL for accessing the file
              uploadedAt: admin.firestore.FieldValue.serverTimestamp(), // Firestore server timestamp
              fileType: file.mimetype, // Store file type for reference
              fileName: sanitizedFileName, // Store sanitized file name
            };

            const mediaRef = await db.collection('gallery').add(mediaData);

            uploadedFiles.push({
              id: mediaRef.id, // Include document ID
              storagePath: mediaData.storagePath,
              url: mediaData.url,
              uploadedAt: mediaData.uploadedAt,
              fileType: mediaData.fileType,
              fileName: mediaData.fileName,
            });
            resolve();
          } catch (err) {
            console.error('❌ Error generating signed URL or storing in Firestore:', err.message);
            reject(err);
          }
        });

        stream.end(file.buffer);
      });
    }

    res.status(200).json({
      message: 'Media uploaded successfully',
      files: uploadedFiles,
    });
  } catch (error) {
    console.error('❌ Server error:', error.message);
    res.status(500).json({ error: `Server error: ${error.message}` });
  }
});

app.get('/gallery', async (req, res) => {
  try {
    const snapshot = await db.collection('gallery').get();
    const images = [];
    console.log("printing snapshotS",snapshot)
    snapshot.forEach(doc => {
      images.push({ id: doc.id, ...doc.data() });
    });
    res.render('gallery', {
      images,
      error: null,
      success: null
    });
    console.log("images",images);
  } catch (error) {
    console.error('❌ Error fetching gallery:', error);
    res.render('gallery', {
      images: [],
      error: 'Failed to load gallery media. Please try again later.',
      success: null
    });
  }
});

app.post('/mark-media-seen/:mediaId', async (req, res) => {
  const { mediaId } = req.params;
  const { seen } = req.body;

  try {
    // Validate request body
    if (typeof seen !== 'boolean') {
      return res.status(400).json({ error: 'Invalid request: seen must be a boolean' });
    }

    // Find and update the media item
    const media = await Media.findOneAndUpdate(
      { mediaId },
      { seen },
      { new: true }
    );

    if (!media) {
      return res.status(404).json({ error: 'Media not found' });
    }

    res.status(200).json({ message: `Media ${mediaId} marked as ${seen ? 'seen' : 'unseen'}`, media });
  } catch (err) {
    console.error('Error updating media:', err);
    res.status(500).json({ error: 'Server error' });
  }
});

app.post('/update-description', requireAdmin, async (req, res) => {
  try {
    const { mediaId, description } = req.body;
    await db.collection('gallery').doc(mediaId).update({ description });
    res.status(200).json({ message: 'Description updated' });
  } catch (error) {
    console.error('❌ Error updating description:', error);
    res.status(500).json({ error: 'Server error' });
  }
});
// School dashboard (renders schoolDashboard.ejs)

// School Dashboard Route
app.get('/school-dashboard', async (req, res) => {
    try {
        const schoolEmail = req.query.username;
        console.log('Request query:', req.query); // Debug log

        // Format time helper
        const formatTime = (timeInput) => {
            if (!timeInput) return null;
            if (timeInput.toDate) {
                const dateObj = timeInput.toDate();
                return dateObj.toLocaleTimeString('en-IN', {
                    hour: '2-digit',
                    minute: '2-digit',
                    hour12: true,
                });
            }
            const timeRegex = /^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/;
            if (typeof timeInput === 'string' && timeRegex.test(timeInput)) {
                const [hours, minutes] = timeInput.split(':').map(Number);
                const date = new Date();
                date.setHours(hours, minutes);
                return date.toLocaleTimeString('en-IN', {
                    hour: '2-digit',
                    minute: '2-digit',
                    hour12: true,
                });
            }
            const dateObj = new Date(timeInput);
            if (!isNaN(dateObj.getTime())) {
                return dateObj.toLocaleTimeString('en-IN', {
                    hour: '2-digit',
                    minute: '2-digit',
                    hour12: true,
                });
            }
            return null;
        };

        // Default response
        const defaultResponse = {
            schoolName: 'Unknown',
            schoolEmail: '',
            city: '',
            district: '',
            pincode: '',
            schoolPhoneNumber: '',
            principalNumber: '',
            principalEmail: '',
            civicsTeacherNumber: '',
            civicsTeacherEmail: '',
            students: [],
            eventDate: null,
            eventDateMissing: true,
            resourcesConfirmed: false,
            selectedResources: [],
            selectedTrainers: [],
            coordinatorDetails: null,
            workshopStartTime: null,
            workshopEndTime: null,
            error: null,
            trialTests: [],
            trainers: [],
            mediaUploads: [],
            totalStudents: 0,
            mcqCompletedCount: 0,
            championCount: 0,
        };

        if (!schoolEmail) {
            console.log('No schoolEmail provided');
            defaultResponse.error = 'Please login first';
            return res.render('schoolDashboard', defaultResponse);
        }

        // Fetch school
        const schoolSnapshot = await db.collection('schools')
            .where('schoolEmail', '==', schoolEmail)
            .get();

        if (schoolSnapshot.empty) {
            console.log('School not found for email:', schoolEmail);
            defaultResponse.error = 'School not found';
            return res.render('schoolDashboard', defaultResponse);
        }

        const schoolData = schoolSnapshot.docs[0].data();
        const schoolId = schoolSnapshot.docs[0].id;
        const schoolName = schoolData.schoolName || 'Unknown';

        // Event Date
        let eventDate = null;
        let eventDateMissing = true;
        if (schoolData.eventDate && typeof schoolData.eventDate.toDate === 'function') {
            eventDate = schoolData.eventDate.toDate().toLocaleDateString('en-US', {
                year: 'numeric',
                month: 'long',
                day: 'numeric',
            });
            eventDateMissing = false;
        } else if (schoolData.eventDate) {
            const parsedDate = new Date(schoolData.eventDate);
            if (!isNaN(parsedDate.getTime())) {
                eventDate = parsedDate.toLocaleDateString('en-US', {
                    year: 'numeric',
                    month: 'long',
                    day: 'numeric',
                });
                eventDateMissing = false;
            }
        }

        // Workshop times
        const workshopStartTime = formatTime(schoolData.workshopStartTime);
        const workshopEndTime = formatTime(schoolData.workshopEndTime);

        // Fetch students
        const studentsSnapshot = await db.collection('participants')
            .where('schoolNameDropdown', '==', schoolName)
            .get();

        const students = studentsSnapshot.docs.map(doc => {
            const data = doc.data();
            return {
                id: doc.id,
                studentName: data.studentName || 'N/A',
                studentClass: data.studentClass || 'N/A',
                hasCompletedMCQ: data.hasCompletedMCQ || false,
                trial1Percentage: data.trial1Percentage || 'N/A',
                trial2Percentage: data.trial2Percentage || 'N/A',
                score: data.score || 0,
                totalQuestions: data.totalQuestions || 0,
                percentage: data.percentage || 0,
                completedAt: data.completedAt ? 
                    (data.completedAt.toDate ? data.completedAt.toDate().toLocaleDateString() : new Date(data.completedAt).toLocaleDateString()) 
                    : 'N/A',
                isApproved: data.isApproved || false,
                isChampion: data.isChampion || false,
                message: data.message || null,
            };
        });

        const totalStudents = students.length;
        const mcqCompletedCount = students.filter(s => s.hasCompletedMCQ).length;
        const championCount = students.filter(s => s.isChampion).length;

        // Fetch trainers
        const trainersSnapshot = await db.collection('trainers').get();
        const trainers = trainersSnapshot.docs.map(doc => {
            const data = doc.data();
            return {
                id: doc.id,
                trainerName: data.trainerName || 'Unnamed',
                district: data.district || 'Not Specified',
                mobileNumber: data.mobileNumber || 'N/A',
                email: data.email || 'N/A',
                availableDates: Array.isArray(data.availableDates)
                    ? data.availableDates.map(d => d.toDate ? d.toDate().toISOString().split('T')[0] : (d.date || d))
                    : [],
            };
        });

        // Fetch coordinator details using coordinatorId from schools collection
        let coordinatorDetails = null;
        if (schoolData.coordinatorId) {
            console.log('Fetching coordinator for ID:', schoolData.coordinatorId);
            const coordinatorDoc = await db.collection('coordinators').doc(schoolData.coordinatorId).get();
            if (coordinatorDoc.exists) {
                const data = coordinatorDoc.data();
                coordinatorDetails = {
                    id: coordinatorDoc.id,
                    name: data.name || 'Unknown', 
                    phone: data.mobile || 'N/A', 
                };
            } else {
                console.log('Coordinator not found for ID:', schoolData.coordinatorId);
            }
        } else {
            console.log('No coordinatorId found for school:', schoolId);
        }

        // Selected trainers with details
        const selectedTrainers = [
            schoolData.trainerId1 || '',
            schoolData.trainerId2 || '',
        ].map(trainerId => {
            const trainer = trainers.find(t => t.id === trainerId);
            return trainer
                ? {
                      ...trainer,
                      isOtherDistrict: trainer.district !== schoolData.district,
                      phone: trainer.mobileNumber || 'N/A',
                  }
                : null;
        }).filter(t => t !== null);

        // Fetch media uploads
        const mediaSnapshot = await db.collection('schools')
            .doc(schoolId)
            .collection('mediaUploads')
            .get();

        const mediaUploads = mediaSnapshot.docs.map(doc => ({
            ...doc.data(),
            id: doc.id,
            uploadedAt: doc.data().uploadedAt?.toDate() || null,
            uploadedBy: schoolName,
            uploaderType: 'School',
        }));

        // Debug log for coordinatorDetails
        console.log('coordinatorDetails before render:', coordinatorDetails);

        // Render dashboard
        res.render('schoolDashboard', {
            schoolName,
            schoolEmail: schoolData.schoolEmail || '',
            city: schoolData.city || '',
            district: schoolData.district || '',
            pincode: schoolData.pincode || '',
            schoolPhoneNumber: schoolData.schoolPhoneNumber || '',
            principalNumber: schoolData.principalNumber || '',
            principalEmail: schoolData.principalEmail || '',
            civicsTeacherNumber: schoolData.civicsTeacherNumber || '',
            civicsTeacherEmail: schoolData.civicsTeacherEmail || '',
            students,
            eventDate,
            eventDateMissing,
            resourcesConfirmed: schoolData.resourcesConfirmed || false,
            selectedResources: schoolData.selectedResources || [],
            selectedTrainers,
            coordinatorDetails,
            workshopStartTime,
            workshopEndTime,
            error: null,
            trialTests: [],
            trainers,
            mediaUploads,
            totalStudents,
            mcqCompletedCount,
            championCount,
        });

    } catch (error) {
        console.error('Error in /school-dashboard route:', error.message, error.stack);
        res.render('schoolDashboard', {
            ...defaultResponse,
            error: 'Error loading school data: ' + error.message,
        });
    }
});


app.post('/school-dashboard/delete-student/:id', async (req, res) => {
    const participantId = req.params.id;
    const schoolEmail = req.body.username || req.query.username;

    try {
        // Delete the student from Firestore
        await db.collection('participants').doc(participantId).delete();
        res.redirect(`/school-dashboard?username=${schoolEmail}&success=Student deleted successfully`);
    } catch (err) {
        console.error('❌ Error deleting student:', err.message);
        res.redirect(`/school-dashboard?username=${schoolEmail}&error=Failed to delete student`);
    }
});
// POST route for submitting resources
    app.post('/school-dashboard/submit-resources', async (req, res) => {
        try {
            console.log("🔧 Incoming form body:", req.body);

            const rawEmail = req.body.schoolEmail;
            if (typeof rawEmail !== 'string') {
                throw new Error("Invalid or missing email.");
            }

            const schoolEmail = rawEmail.trim().toLowerCase();
            const { resources, trainerId1, trainerId2 ,workshopStartTime, workshopEndTime} = req.body;

            console.log("📨 Cleaned Email:", schoolEmail);

            // 🔍 Step 1: Try normal Firestore query
            let schoolDoc = null;
            const snapshot = await db.collection('schools')
                .where('schoolEmail', '==', schoolEmail)
                .get();

            if (!snapshot.empty) {
                schoolDoc = snapshot.docs[0];
            }

            // 🔍 Step 2: Fallback if needed (case/space mismatch)
            if (!schoolDoc) {
                console.log("❌ Exact match failed. Trying fallback...");
                const allSchools = await db.collection('schools').get();
                allSchools.forEach(doc => {
                    const data = doc.data();
                    if (
                        data.schoolEmail &&
                        data.schoolEmail.trim().toLowerCase() === schoolEmail
                    ) {
                        schoolDoc = doc;
                    }
                });
            }

            // ❌ Step 3: If still not found
            if (!schoolDoc) {
                console.log("❌ School not found after fallback.");

                return res.status(404).render('schoolDashboard', {
                    schoolName: '',
                    schoolEmail,
                    city: '',
                    district: '',
                    pincode: '',
                    schoolPhoneNumber: '',
                    principalNumber: '',
                    principalEmail: '',
                    civicsTeacherNumber: '',
                    civicsTeacherEmail: '',
                    students: [],
                    eventDate: null,
                    eventDateMissing: true,
                    resourcesConfirmed: false,
                    selectedResources: [],
                    selectedTrainers: [],
                    trainers: [],
                     workshopStartTime: null,
                    workshopEndTime: null,
                    mediaUploads: [],
                    trialTests,
                    error: 'School not found.'
                });
            }

            // ✅ Step 4: School found
            const schoolData = schoolDoc.data();
            const schoolId = schoolDoc.id;

            const selectedResources = Array.isArray(resources)
                ? resources
                : [resources];

            // 🎯 Step 5: Fetch trainer details with transformed availableDates
            const trainer1Doc = await db.collection('trainers').doc(trainerId1).get();
            const trainer2Doc = await db.collection('trainers').doc(trainerId2).get();

            const selectedTrainers = [];

            if (trainer1Doc.exists) {
                const trainer = trainer1Doc.data();
                selectedTrainers.push({
                    ...trainer,
                    isOtherDistrict: trainer.district !== schoolData.district,
                    availableDates: Array.isArray(trainer.availableDates) ? trainer.availableDates.map(dateObj => 
                        dateObj.date || dateObj.toDate ? dateObj.toDate().toISOString().split('T')[0] : 'N/A'
                    ) : []
                });
            }

            if (trainer2Doc.exists) {
                const trainer = trainer2Doc.data();
                selectedTrainers.push({
                    ...trainer,
                    isOtherDistrict: trainer.district !== schoolData.district,
                    availableDates: Array.isArray(trainer.availableDates) ? trainer.availableDates.map(dateObj => 
                        dateObj.date || dateObj.toDate ? dateObj.toDate().toISOString().split('T')[0] : 'N/A'
                    ) : []
                });
            }

            // 📝 Step 6: Update school data
            await db.collection('schools').doc(schoolId).update({
                resourcesConfirmed: true,
                selectedResources,
                 workshopStartTime: workshopStartTime || '',
                workshopEndTime: workshopEndTime || '',
                trainerId1,
                trainerId2
            });

            // 📚 Step 7: Load all trainers for dropdown
            const allTrainers = await db.collection('trainers').get();
            const trainers = allTrainers.docs.map(doc => {
                const data = doc.data();
                return {
                    id: doc.id,
                    ...data,
                    availableDates: Array.isArray(data.availableDates) ? data.availableDates.map(dateObj => 
                        dateObj.date || dateObj.toDate ? dateObj.toDate().toISOString().split('T')[0] : 'N/A'
                    ) : []
                };
            });

            // ✅ Step 8: Render school dashboard with confirmation
            return res.render('schoolDashboard', {
                schoolName: schoolData.schoolName || '',
                schoolEmail,
                city: schoolData.city || '',
                district: schoolData.district || '',
                pincode: schoolData.pincode || '',
                schoolPhoneNumber: schoolData.schoolPhoneNumber || '',
                principalNumber: schoolData.principalNumber || '',
                principalEmail: schoolData.principalEmail || '',
                civicsTeacherNumber: schoolData.civicsTeacherNumber || '',
                civicsTeacherEmail: schoolData.civicsTeacherEmail || '',
                students: schoolData.students || [],
                eventDate: schoolData.eventDate ? schoolData.eventDate.toDate().toLocaleDateString('en-US', {
                    year: 'numeric',
                    month: 'long',
                    day: 'numeric'
                }) : null,
                eventDateMissing: !schoolData.eventDate,
                resourcesConfirmed: true,
                selectedResources,
                selectedTrainers,
                 workshopStartTime: workshopStartTime || null,
                workshopEndTime: workshopEndTime || null,
                trainers,
                mediaUploads: [], // optional: pass uploads if you have them
                trialTests,
                error: null
            });

        } catch (error) {
            console.error("❌ Error in /submit-resources:", error);

            res.status(500).render('schoolDashboard', {
                schoolName: '',
                schoolEmail: req.body?.schoolEmail || '',
                city: '',
                district: '',
                pincode: '',
                schoolPhoneNumber: '',
                principalNumber: '',
                principalEmail: '',
                civicsTeacherNumber: '',
                civicsTeacherEmail: '',
                students: [],
                eventDate: null,
                eventDateMissing: true,
                resourcesConfirmed: false,
                selectedResources: [],
                selectedTrainers: [],
                trainers: [],
                mediaUploads: [],
                  workshopStartTime: null,
                workshopEndTime: null,
                trialTests,
                error: 'Something went wrong while submitting resources.'
            });
        }
    });

    // Update school information
    app.post('/school-dashboard/update', [
        body('city').trim().notEmpty().withMessage('City is required'),
        body('district').trim().notEmpty().withMessage('District is required'),
        body('pincode').trim().matches(/^\d{6}$/).withMessage('Pincode must be 6 digits'),
        body('schoolPhoneNumber').trim().matches(/^\d{10}$/).withMessage('School phone number must be 10 digits'),
        body('principalNumber').trim().matches(/^\d{10}$/).withMessage('Principal phone number must be 10 digits'),
        body('principalEmail').trim().isEmail().withMessage('Invalid principal email address'),
        body('civicsTeacherNumber').trim().matches(/^\d{10}$/).withMessage('Civics teacher phone number must be 10 digits'),
        body('civicsTeacherEmail').trim().isEmail().withMessage('Invalid civics teacher email address')
    ], async (req, res) => {
        try {
            const errors = validationResult(req);
            if (!errors.isEmpty()) {
                const schoolEmail = req.body.schoolEmail;
                return res.status(400).render('schoolDashboard', {
                    schoolName: 'Unknown',
                    schoolEmail,
                    city: req.body.city,
                    district: req.body.district,
                    pincode: req.body.pincode,
                    schoolPhoneNumber: req.body.schoolPhoneNumber,
                    principalNumber: req.body.principalNumber,
                    principalEmail: req.body.principalEmail,
                    civicsTeacherNumber: req.body.civicsTeacherNumber,
                    civicsTeacherEmail: req.body.civicsTeacherEmail,
                    students: [],
                    eventDate: null,
                    eventDateMissing: true,
                    resourcesConfirmed: false,
                    selectedResources: [],
                    error: errors.array()[0].msg,
                    trialTests
                });
            }

            const {
                schoolEmail,
                city,
                district,
                pincode,
                schoolPhoneNumber,
                principalNumber,
                principalEmail,
                civicsTeacherNumber,
                civicsTeacherEmail
            } = req.body;

            const schoolSnapshot = await db.collection('schools')
                .where('schoolEmail', '==', schoolEmail)
                .get();
            if (schoolSnapshot.empty) {
                return res.status(404).render('schoolDashboard', {
                    schoolName: 'Unknown',
                    schoolEmail,
                    city,
                    district,
                    pincode,
                    schoolPhoneNumber,
                    principalNumber,
                    principalEmail,
                    civicsTeacherNumber,
                    civicsTeacherEmail,
                    students: [],
                    eventDate: null,
                    eventDateMissing: true,
                    resourcesConfirmed: false,
                    selectedResources: [],
                    error: 'School not found',
                    trialTests
                });
            }

            const schoolId = schoolSnapshot.docs[0].id;
            await db.collection('schools').doc(schoolId).update({
                city,
                district,
                pincode,
                schoolPhoneNumber,
                principalNumber,
                principalEmail,
                civicsTeacherNumber,
                civicsTeacherEmail
            });

            res.redirect(`/school-dashboard?username=${encodeURIComponent(schoolEmail)}`);
        } catch (error) {
            console.error('Error in school-dashboard/update route:', error.message, error.stack);
            res.status(500).render('schoolDashboard', {
                schoolName: 'Unknown',
                schoolEmail: req.body.schoolEmail,
                city: req.body.city,
                district: req.body.district,
                pincode: req.body.pincode,
                schoolPhoneNumber: req.body.schoolPhoneNumber,
                principalNumber: req.body.principalNumber,
                principalEmail: req.body.principalEmail,
                civicsTeacherNumber: req.body.civicsTeacherNumber,
                civicsTeacherEmail: req.body.civicsTeacherEmail,
                students: [],
                eventDate: null,
                eventDateMissing: true,
                resourcesConfirmed: false,
                selectedResources: [],
                error: 'Error updating school information',
                trialTests
            });
        }
    });

    // School login (renders login.ejs)
    app.post('/school-login', async (req, res) => {
        const { username, password } = req.body;
        console.log(`School login attempt: username=${username}`); // Add logging
        try {
            if (!username || !password) {
                console.log('Missing username or password');
                return res.render('login', { error: 'Username and password are required.' });
            }

            const snapshot = await db.collection('schools')
                .where('schoolEmail', '==', username)
                .get();
            if (snapshot.empty) {
                console.log(`No school found for username: ${username}`);
                return res.render('login', { error: 'Invalid username or password.' });
            }

            const school = snapshot.docs[0].data();
            if (password !== school.principalNumber) {

                console.log(`Invalid password for username: ${username} , ${school.principalNumber}, ${password}`);
                return res.render('login', { error: 'Invalid username or password.' });
            }

            console.log(`Login successful for ${username}, redirecting to school-dashboard`);
            // Store schoolEmail in session for consistency
            req.session.schoolEmail = username;
            res.redirect(`/school-dashboard?username=${encodeURIComponent(username)}`);
        } catch (error) {
            console.error('Error during school login:', error.message, error.stack);
            res.render('login', { error: 'Login failed. Try again later.' });
        }
    });

    // School upload media 
    app.post('/school-dashboard/upload-media', uploadSchoolMedia , async (req, res) => {
        try {
            const { schoolEmail, mediaDescription, mediaLink } = req.body;
            const photos = req.files['photos'] || [];
            const videos = req.files['videos'] || [];

            if (!schoolEmail) {
                return res.status(400).render('schoolDashboard', {
                    schoolName: 'Unknown',
                    schoolEmail: '',
                    city: '',
                    district: '',
                    pincode: '',
                    schoolPhoneNumber: '',
                    principalNumber: '',
                    principalEmail: '',
                    civicsTeacherNumber: '',
                    civicsTeacherEmail: '',
                    students: [],
                    eventDate: null,
                    eventDateMissing: true,
                    resourcesConfirmed: false,
                    selectedResources: [],
                    selectedTrainers: [],
                    trainers: [],
                    mediaUploads: [],
                    error: 'School email is required',
                    trialTests
                });
            }

            const schoolSnapshot = await db.collection('schools')
                .where('schoolEmail', '==', schoolEmail)
                .get();
            if (schoolSnapshot.empty) {
                return res.status(404).render('schoolDashboard', {
                    schoolName: 'Unknown',
                    schoolEmail,
                    city: '',
                    district: '',
                    pincode: '',
                    schoolPhoneNumber: '',
                    principalNumber: '',
                    principalEmail: '',
                    civicsTeacherNumber: '',
                    civicsTeacherEmail: '',
                    students: [],
                    eventDate: null,
                    eventDateMissing: true,
                    resourcesConfirmed: false,
                    selectedResources: [],
                    selectedTrainers: [],
                    trainers: [],
                    mediaUploads: [],
                    error: 'School not found',
                    trialTests
                });
            }

            const schoolId = schoolSnapshot.docs[0].id;
            const schoolData = schoolSnapshot.docs[0].data();

            if (mediaDescription && mediaDescription.length > 1000) {
                return res.status(400).render('schoolDashboard', {
                    schoolName: schoolData.schoolName,
                    schoolEmail,
                    city: schoolData.city || '',
                    district: schoolData.district || schoolData.city || '',
                    pincode: schoolData.pincode || '',
                    schoolPhoneNumber: schoolData.schoolPhoneNumber || '',
                    principalNumber: schoolData.principalNumber || '',
                    principalEmail: schoolData.principalEmail || '',
                    civicsTeacherNumber: schoolData.civicsTeacherNumber || '',
                    civicsTeacherEmail: schoolData.civicsTeacherEmail || '',
                    students: [],
                    eventDate: schoolData.eventDate ? schoolData.eventDate.toDate().toLocaleDateString('en-US') : null,
                    eventDateMissing: !schoolData.eventDate,
                    resourcesConfirmed: schoolData.resourcesConfirmed || false,
                    selectedResources: schoolData.selectedResources || [],
                    selectedTrainers: [
                        schoolData.selectedTrainer1 || '',
                        schoolData.selectedTrainer2 || ''
                    ],
                    trainers: [],
                    mediaUploads: [],
                    error: 'Media description cannot exceed 1000 words',
                    trialTests
                });
            }

            // Explicitly set the bucket name
            const bucket = admin.storage().bucket('beinglawful-ee5a4.firebasestorage.app'); // Verify this matches your bucket
            const [exists] = await bucket.exists();
            if (!exists) {
                throw new Error('The specified bucket does not exist. Please check your Firebase configuration.');
            }

            const mediaUploads = [];
            for (const file of photos) {
                const fileName = `media/${schoolId}/photos/${Date.now()}_${file.originalname}`;
                const fileUpload = bucket.file(fileName);
                try {
                    await fileUpload.save(file.buffer, {
                        metadata: { contentType: file.mimetype }
                    });
                    const [url] = await fileUpload.getSignedUrl({
                        action: 'read',
                        expires: '03-09-2491'
                    });
                    mediaUploads.push({
                        url,
                        type: 'image',
                        path: fileName,
                        description: mediaDescription || '',
                        link: mediaLink || '',
                        uploadedAt: admin.firestore.FieldValue.serverTimestamp()
                    });
                } catch (uploadError) {
                    console.error(`Error uploading photo ${fileName}:`, uploadError.message, uploadError.stack);
                    throw new Error(`Failed to upload photo: ${uploadError.message}`);
                }
            }

            for (const file of videos) {
                const fileName = `media/${schoolId}/videos/${Date.now()}_${file.originalname}`;
                const fileUpload = bucket.file(fileName);
                try {
                    await fileUpload.save(file.buffer, {
                        metadata: { contentType: file.mimetype }
                    });
                    const [url] = await fileUpload.getSignedUrl({
                        action: 'read',
                        expires: '03-09-2491'
                    });
                    mediaUploads.push({
                        url,
                        type: 'video',
                        path: fileName,
                        description: mediaDescription || '',
                        link: mediaLink || '',
                        uploadedAt: admin.firestore.FieldValue.serverTimestamp()
                    });
                } catch (uploadError) {
                    console.error(`Error uploading video ${fileName}:`, uploadError.message, uploadError.stack);
                    throw new Error(`Failed to upload video: ${uploadError.message}`);
                }
            }

            if (mediaUploads.length > 0) {
                const batch = db.batch();
                mediaUploads.forEach(upload => {
                    const mediaRef = db.collection('schools').doc(schoolId).collection('mediaUploads').doc();
                    batch.set(mediaRef, upload);
                });
                await batch.commit();
            }

            const trainersSnapshot = await db.collection('trainers').get();
            const trainers = trainersSnapshot.docs.map(doc => ({
                id: doc.id,
                trainerName: doc.data().trainerName || 'Unnamed',
                district: doc.data().district || 'Not Specified'
            }));

            const updatedSchoolSnapshot = await db.collection('schools').doc(schoolId).get();
            const updatedSchoolData = updatedSchoolSnapshot.data();

            res.render('schoolDashboard', {
                schoolName: updatedSchoolData.schoolName || '',
                schoolEmail,
                city: updatedSchoolData.city || '',
                district: updatedSchoolData.district || updatedSchoolData.city || '',
                pincode: updatedSchoolData.pincode || '',
                schoolPhoneNumber: updatedSchoolData.schoolPhoneNumber || '',
                principalNumber: updatedSchoolData.principalNumber || '',
                principalEmail: updatedSchoolData.principalEmail || '',
                civicsTeacherNumber: updatedSchoolData.civicsTeacherNumber || '',
                civicsTeacherEmail: updatedSchoolData.civicsTeacherEmail || '',
                students: [], // Fetch students if needed
                eventDate: updatedSchoolData.eventDate ? updatedSchoolData.eventDate.toDate().toLocaleDateString('en-US') : null,
                eventDateMissing: !updatedSchoolData.eventDate,
                resourcesConfirmed: updatedSchoolData.resourcesConfirmed || false,
                selectedResources: updatedSchoolData.selectedResources || [],
                selectedTrainers: [
                    updatedSchoolData.selectedTrainer1 || '',
                    updatedSchoolData.selectedTrainer2 || ''
                ],
                trainers,
                mediaUploads,
                error: null,
                trialTests
            });
        } catch (error) {
            console.error('Error in school-dashboard/upload-media route:', error.message, error.stack);
            const schoolEmail = req.body.schoolEmail || '';
            res.status(500).render('schoolDashboard', {
                schoolName: 'Unknown',
                schoolEmail,
                city: '',
                district: '',
                pincode: '',
                schoolPhoneNumber: '',
                principalNumber: '',
                principalEmail: '',
                civicsTeacherNumber: '',
                civicsTeacherEmail: '',
                students: [],
                eventDate: null,
                eventDateMissing: true,
                resourcesConfirmed: false,
                selectedResources: [],
                selectedTrainers: [],
                trainers: [],
                mediaUploads: [],
                error: `Error uploading media: ${error.message}`,
                trialTests
            });
        }
    });
    



    // admin to fetch photos and videos from Firestore
// Route to fetch photos and videos from Firestore
    app.get('/fetch-media', requireAdmin, async (req, res) => {
    try {
        const { parentMobile1, trainerEmail, schoolEmail } = req.query; // Optional filters
        const imageUploads = {
        images: {
            school: [],
            trainer: [],
            student: []
        },
        videos: {
            school: [],
            trainer: [],
            student: []
        },
        all: [] // For sorting all media by uploadedAt
        };

        // Helper function to fetch media from a collection
        async function fetchMediaFromCollection(collectionName, identifierField, identifierValue, uploaderType, nameField) {
            console.log
        let query = db.collection(collectionName);
        if (identifierValue) {
            query = query.where(identifierField, '==', identifierValue);
        }

        const snapshot = await query.get();
        if (snapshot.empty) {
            console.warn(`No ${collectionName} found for ${identifierField}: ${identifierValue || 'all'}`);
            return;
        }

        for (const doc of snapshot.docs) {
            const uploaderName = doc.data()[nameField] || 'Unknown';
            const mediaSnapshot = await doc.ref.collection('mediaUploads').get();
            mediaSnapshot.forEach(mediaDoc => {
            const mediaData = mediaDoc.data();
            const mediaItem = {
                id: mediaDoc.id,
                url: mediaData.url,
                path: mediaData.path,
                type: mediaData.type,
                description: mediaData.description || '',
                link: mediaData.link || '',
                uploadedAt: mediaData.uploadedAt ? mediaData.uploadedAt.toDate() : null,
                ownerId: doc.id,
                uploaderType,
                seen: mediaData.seen || false,
                uploadedBy: uploaderName
            };

            if (mediaData.type === 'image') {
                imageUploads.images[uploaderType].push(mediaItem);
            } else if (mediaData.type === 'video') {
                imageUploads.videos[uploaderType].push(mediaItem);
            }
            imageUploads.all.push(mediaItem);
                       });
        }
        }

        // Fetch media from all collections
        await Promise.all([
        fetchMediaFromCollection('participants', 'parentMobile1', parentMobile1, 'student', 'studentName'),
        fetchMediaFromCollection('trainers', 'email', trainerEmail, 'trainer', 'trainerName'),
        fetchMediaFromCollection('schools', 'schoolEmail', schoolEmail, 'school', 'schoolName')
        ]);

        // Sort all media by uploadedAt (descending)
        imageUploads.all.sort((a, b) => (b.uploadedAt || 0) - (a.uploadedAt || 0));
            
        res.render('media', {
        imageUploads,
        error: null,
        success: req.query.success || null
        });
    } catch (error) {
        console.error('Error fetching media from Firestore:', error.message, error.stack);
        res.render('media', {
        imageUploads: {
            images: { school: [], trainer: [], student: [] },
            videos: { school: [], trainer: [], student: [] },
            all: []
        },
        error: `Error fetching media: ${error.message}`,
        success: null
        });
    }
    });

app.post('/school-dashboard/approve-student/:id', async (req, res) => {
    try {
        const studentId = req.params.id;
        const schoolEmail = req.query.username;
        if (!studentId || !schoolEmail) return res.status(400).send('Missing required fields');

        await db.collection('participants').doc(studentId).update({ isApproved: true });

        res.redirect(`/school-dashboard?username=${encodeURIComponent(schoolEmail)}`);
    } catch (err) {
        console.error('Error approving student:', err);
        res.status(500).send('Error approving student.');
    }
});

//champ student
 app.post('/school-dashboard/mark-champ/:id', async (req, res) => {
    const studentId = req.params.id;
    const { username } = req.query; // This is the schoolEmail

    try {
        // ✅ Count champions for this school only
        const championsSnapshot = await db.collection('participants')
            .where('isChampion', '==', true)
            .where('schoolEmail', '==', username)
            .get();

        if (championsSnapshot.size >= 50) {
            return res.send('⚠️ Only 50 students can be marked as Champion for your school.');
        }

        const championMessage = '🎉 Congratulations! You have been selected as a Being Lawful Champ!';

        // ✅ Mark this student as Champion
        await db.collection('participants').doc(studentId).update({
            isChampion: true,
            championMessage: championMessage
        });

        console.log(`Champion set for ${studentId} - Message: ${championMessage}`);

        res.redirect(`/school-dashboard?username=${username}`);
    } catch (error) {
        console.error('Error marking champ:', error);
        res.status(500).send('Server error while marking champion');
    }
});

// Participation form (renders participation.ejs)
    app.get('/participation', async (req, res) => {
        try {
            const type = req.query.type || 'Student';

            const snapshot = await db.collection('schools')
                .where('isApproved', '==', true)
                .get();

            const schoolNames = snapshot.docs.map(doc => doc.data().schoolName);

            // ✅ Pass errors and values even if empty
            res.render('participation', {
                type,
                schoolNames,
                errors: {},  // ← Prevents ReferenceError
                values: {}   // ← Keeps fields empty on first load
            });
        } catch (error) {
            console.error('Error in participation route:', error.message, error.stack);
            res.status(500).send('Error loading participation form.');
        }
    });

   app.post('/participate', async (req, res) => {
    try {
        const participant = {
            studentName: req.body.studentName,
            schoolNameDropdown: req.body.schoolNameDropdown,
            birthdate: req.body.birthdate,
            studentClass: req.body.studentClass,
            parentMobile1: req.body.parentMobile1,
            parentMobile2: req.body.parentMobile2,
            parentEmail: req.body.parentEmail,
            address: req.body.address,
            city: req.body.city,
            pincode: req.body.pincode,
            type: req.body.type,
        };

        const errors = {};

        // 🔍 Find schoolEmail using selected school name
        const schoolSnapshot = await db.collection('schools')
            .where('schoolName', '==', participant.schoolNameDropdown)
            .limit(1)
            .get();

        if (!schoolSnapshot.empty) {
            const schoolData = schoolSnapshot.docs[0].data();
            participant.schoolEmail = schoolData.email || '';
        } else {
            errors.schoolNameDropdown = 'Selected school not found';
        }

        // ✅ Validation checks
        const emailSnapshot = await db.collection('participants')
            .where('parentEmail', '==', participant.parentEmail)
            .get();
        if (!emailSnapshot.empty) {
            errors.parentEmail = 'User with this email already exists';
        }

        const mobileSnapshot = await db.collection('participants')
            .where('parentMobile1', '==', participant.parentMobile1)
            .get();
        if (!mobileSnapshot.empty) {
            errors.parentMobile1 = 'User with this mobile number already exists';
        }

        const requiredFields = ['studentName', 'schoolNameDropdown', 'birthdate', 'studentClass', 'parentMobile1', 'parentEmail', 'address', 'city', 'pincode', 'type'];
        requiredFields.forEach(field => {
            if (!participant[field]) {
                errors[field] = 'This field is required';
            }
        });

        if (participant.pincode && !/^\d{6}$/.test(participant.pincode)) {
            errors.pincode = 'Pincode must be 6 digits';
        }

        if (participant.parentMobile1 && !/^\d{10}$/.test(participant.parentMobile1)) {
            errors.parentMobile1 = 'Mobile number must be 10 digits';
        }

        if (participant.parentMobile2 && participant.parentMobile2 !== '' && !/^\d{10}$/.test(participant.parentMobile2)) {
            errors.parentMobile2 = 'Mobile number must be 10 digits';
        }

        if (participant.parentEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(participant.parentEmail)) {
            errors.parentEmail = 'Invalid email address';
        }

        if (participant.birthdate && !/^\d{4}-\d{2}-\d{2}$/.test(participant.birthdate)) {
            errors.birthdate = 'Birthdate must be YYYY-MM-DD';
        }

        // ❌ If any errors, re-render form
        if (Object.keys(errors).length > 0) {
            const snapshot = await db.collection('schools')
                .where('isApproved', '==', true)
                .get();
            const schoolNames = snapshot.docs.map(doc => doc.data().schoolName);

            return res.render('participation', {
                type: participant.type || 'Student',
                schoolNames,
                errors,
                values: participant
            });
        }

        // ✅ Save participant with schoolEmail
        await db.collection('participants').add({
            ...participant,
            hasCompletedTrial1: false,
            hasCompletedTrial2: false,
            hasCompletedMCQ: false,
            trainerApprove: false,
            score: 0
        });

        const [year, month, day] = participant.birthdate.split('-');
        const password = `${day}${month}${year}`;

        res.render('studentConfirmation', {
            studentName: participant.studentName,
            username: participant.parentMobile1,
            password
        });

    } catch (error) {
        console.error('Error in participate route:', error.message, error.stack);
        res.status(500).render('participation', {
            type: 'Student',
            schoolNames: [],
            errors: { server: 'Something went wrong. Please try again.' },
            values: req.body
        });
    }
});


// Route to mark a school as completed

// Existing endpoint for completing a school
app.post('/admin/complete-school/:schoolId', async (req, res) => {
    try {
        const { schoolId } = req.params;

        // Validate schoolId format
        if (!schoolId || typeof schoolId !== 'string' || schoolId.trim() === '') {
            return res.status(400).json({ error: 'Invalid school ID' });
        }

        // Check if session exists and user is admin
        if (!req.session || !req.session.isAdmin) {
            return res.status(401).json({ error: 'Unauthorized: Admin access required' });
        }

        // Reference to the school document
        const schoolRef = db.collection('schools').doc(schoolId);
        const schoolDoc = await schoolRef.get();

        // Check if school exists
        if (!schoolDoc.exists) {
            return res.status(404).json({ error: 'School not found' });
        }

        // Check if school is already completed
        const schoolData = schoolDoc.data();
        if (schoolData.isCompleted) {
            return res.status(400).json({ error: 'School is already completed' });
        }

        // Update school document
        await schoolRef.update({
            isCompleted: true,
            completedAt: admin.firestore.FieldValue.serverTimestamp(),
            updatedBy: req.session.userId || 'admin'
        });

        // Fetch updated document
        const updatedSchoolDoc = await schoolRef.get();
        const updatedSchoolData = updatedSchoolDoc.data();

        // Respond with success message
        res.status(200).json({
            message: 'School marked as completed',
            schoolId,
            school: {
                id: schoolId,
                ...updatedSchoolData
            }
        });
    } catch (error) {
        console.error(`Error in complete-school route for schoolId ${schoolId}:`, {
            message: error.message,
            stack: error.stack,
            code: error.code
        });
        res.status(500).json({
            error: 'Server error',
            details: error.message || 'An unexpected error occurred'
        });
    }
});

    // School participation form (renders schoolParticipation.ejs)
    app.get('/school-participation', async (req, res) => {
        try {
            const snapshot = await db.collection('schools').get();
            const schools = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
            res.render('schoolParticipation', { schools, errors: null });
        } catch (error) {
            console.error('Error in school-participation route:', error.message, error.stack);
            res.status(500).send('Error fetching school data.');
        }
    });

   app.post('/school-participate', [
        body('schoolName').trim().notEmpty().withMessage('School name is required'),
        body('city').trim().notEmpty().withMessage('City is required'),
        body('district').trim().notEmpty().withMessage('District is required'),
        body('pincode').trim().matches(/^\d{6}$/).withMessage('Pincode must be 6 digits'),
        body('schoolPhoneNumber').trim().matches(/^\d{10}$/).withMessage('School phone number must be 10 digits'),
        body('schoolEmail').trim().isEmail().withMessage('Invalid school email address'),
        body('principalNumber').trim().matches(/^\d{10}$/).withMessage('Principal phone number must be 10 digits'),
        body('principalEmail').trim().isEmail().withMessage('Invalid principal email address'),
        body('civicsTeacherNumber').trim().matches(/^\d{10}$/).withMessage('Civics teacher phone number must be 10 digits'),
        body('civicsTeacherEmail').trim().isEmail().withMessage('Invalid civics teacher email address'),
        body('eventDate1').isDate().withMessage('Invalid date format for Event Date 1'),
        body('eventDate2').isDate().withMessage('Invalid date format for Event Date 2'),
        body('eventDate3').isDate().withMessage('Invalid date format for Event Date 3'),
        body('eventDate4').isDate().withMessage('Invalid date format for Event Date 4'),
        body('referenceSource').trim().notEmpty().withMessage('Reference source is required'),
        body('referenceName').if(body('referenceSource').equals('Person')).trim().notEmpty().withMessage('Reference name is required when reference source is Person')
    ], async (req, res) => {
        try {
            const errors = validationResult(req).array();

            // Check for duplicate entries
            const {
                schoolName,
                city,
                district,
                pincode,
                schoolPhoneNumber,
                schoolEmail,
                principalNumber,
                principalEmail,
                civicsTeacherNumber,
                civicsTeacherEmail,
                eventDate1,
                eventDate2,
                eventDate3,
                eventDate4,
                referenceSource,
                referenceName
            } = req.body;

            const checks = [
                { field: 'schoolEmail', value: schoolEmail, message: 'School email already exists' },
                { field: 'principalEmail', value: principalEmail, message: 'Principal email already exists' },
                { field: 'civicsTeacherEmail', value: civicsTeacherEmail, message: 'Civics teacher email already exists' },
                { field: 'schoolPhoneNumber', value: schoolPhoneNumber, message: 'School phone number already exists' },
                { field: 'principalNumber', value: principalNumber, message: 'Principal phone number already exists' },
                { field: 'civicsTeacherNumber', value: civicsTeacherNumber, message: 'Civics teacher phone number already exists' }
            ];

            for (const check of checks) {
                const snapshot = await db.collection('schools')
                    .where(check.field, '==', check.value)
                    .get();
                if (!snapshot.empty) {
                    errors.push({ param: check.field, msg: check.message });
                }
            }

            if (errors.length > 0) {
                const schoolsSnapshot = await db.collection('schools').get();
                const schools = schoolsSnapshot.docs.map(doc => {
                    const data = doc.data();
                    return {
                        id: doc.id,
                        ...data,
                        // Convert Firestore Timestamp to readable string
                        registeredAt: data.registeredAt ? data.registeredAt.toDate().toLocaleString() : 'Not registered'
                    };
                });
                return res.status(400).render('schoolParticipation', { schools, errors });
            }

            await db.collection('schools').add({
                schoolName,
                city,
                district,
                pincode,
                schoolPhoneNumber,
                schoolEmail,
                principalNumber,
                principalEmail,
                civicsTeacherNumber,
                civicsTeacherEmail,
                eventDates: [eventDate1, eventDate2, eventDate3, eventDate4],
                referenceSource,
                referenceName: referenceSource === 'Person' ? referenceName : null,
                registeredAt: admin.firestore.FieldValue.serverTimestamp(),
                isApproved: false,
                resourcesConfirmed: false,
                selectedResources: []
            });

            // await sendEmail(
            //     schoolEmail,
            //     emailTemplates.schoolRegistration.subject,
            //     emailTemplates.schoolRegistration.text(schoolName, schoolEmail, principalNumber),
            //     emailTemplates.schoolRegistration.html(schoolName, schoolEmail, principalNumber)
            // );

            res.render('confirmation', {
                schoolEmail: schoolEmail,
                principalNumber: principalNumber
            });
        } catch (error) {
            console.error('Error in school-participate route:', error.message, error.stack);
            const schoolsSnapshot = await db.collection('schools').get();
            const schools = schoolsSnapshot.docs.map(doc => {
                const data = doc.data();
                return {
                    id: doc.id,
                    ...data,
                    // Convert Firestore Timestamp to readable string
                    registeredAt: data.registeredAt ? data.registeredAt.toDate().toLocaleString() : 'Not registered'
                };
            });
            res.status(500).render('schoolParticipation', {
                schools,
                errors: [{ msg: 'Internal server error. Please try again later.' }]
            });
        }
    });
        // Confirmation page (renders confirmation.ejs)
        app.get('/confirmation', (req, res) => {
            res.render('confirmation', {
                schoolEmail: 'Not provided',
                principalNumber: 'Not provided'
            });
        });

        // School students (renders schoolStudents.ejs)
        app.get('/school-students', async (req, res) => {
            try {
                const schoolName = req.query.schoolName;
                if (!schoolName) {
                    return res.status(400).send('School name is required.');
                }

                const snapshot = await db.collection('participants')
                    .where('schoolName', '==', schoolName)
                    .get();
                const students = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));

                res.render('schoolStudents', { schoolName, students });
            } catch (error) {
                console.error('Error in school-students route:', error.message, error.stack);
                res.status(500).send('Error fetching student data.');
            }
        });

        // Admin login page (renders adminLogin.ejs)
        app.get('/admin-login', (req, res) => {
            if (req.session.isAdmin) {
                return res.redirect('/admin-dashboard');
            }
            res.render('adminLogin', { error: null });
        });

        // Admin login (renders adminLogin.ejs on error)
        app.post('/admin-login', (req, res) => {
            const { username, password } = req.body;
            if (!username || !password) {
                return res.render('adminLogin', { error: 'Username and password are required.' });
            }

            if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
                req.session.isAdmin = true;
                res.redirect('/admin-dashboard');
            } else {
                res.render('adminLogin', { error: 'Invalid username or password.' });
            }
        });

        // Admin route (redirects to /admin-dashboard for backward compatibility)
        app.get('/admin', requireAdmin, async (req, res) => {
            res.redirect('/admin-dashboard');
        });





// Helper function to format time

// Helper function to create a worksheet from data
function formatTimestamp(date) {
  if (!date) return '';
  return date.toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
}

function createWorksheet(data, columns) {
  const rows = [];
  // Add header row
  rows.push(columns.map(col => col.header));

  // Add data rows
  for (const item of data) {
    const row = columns.map(col => {
      let val = item[col.key];
      if (col.transform) val = col.transform(val, item);
      return val === undefined || val === null ? '' : val;
    });
    rows.push(row);
  }

  return XLSX.utils.aoa_to_sheet(rows);
}


// Function to format timestamps consistently with frontend (Asia/Kolkata timezone)
function formatTimestamp(date) {
    if (!date) return 'N/A';
    return new Date(date).toLocaleString('en-US', {
        timeZone: 'Asia/Kolkata',
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });
}

// Function to create a worksheet from data and columns
function createWorksheet(data, columns) {
    const ws = XLSX.utils.json_to_sheet(
        data.map(item =>
            columns.reduce((acc, col) => {
                let value = item[col.key];
                if (col.transform) {
                    value = col.transform(value, item);
                }
                acc[col.header] = value ?? 'N/A';
                return acc;
            }, {})
        )
    );
    return ws;
}

app.get('/admin/export-data', requireAdmin, async (req, res) => {
    try {
        console.log('Starting export-data endpoint');
        const type = req.query.type || 'all';
        const workbook = XLSX.utils.book_new();

        let schoolsSnapshot = null; // Initialize to avoid undefined errors
        let trainersSnapshot = null; // Initialize to avoid undefined errors

        if (type === 'schools' || type === 'all') {
            // Fetch Schools
            console.log('Fetching schools...');
            schoolsSnapshot = await db.collection('schools').get();
            const schools = [];
            for (const doc of schoolsSnapshot.docs) {
                const schoolData = doc.data();
                console.log(`Processing school: ${doc.id}`);
                const eventDate = schoolData.eventDate?.toDate ? schoolData.eventDate.toDate() : (schoolData.eventDate ? new Date(schoolData.eventDate) : null);
                let trainer1 = null, trainer2 = null, coordinator = null;
                if (schoolData.trainerId1) {
                    const t1Doc = await db.collection('trainers').doc(schoolData.trainerId1).get();
                    if (t1Doc.exists) trainer1 = { id: t1Doc.id, ...t1Doc.data() };
                }
                if (schoolData.trainerId2) {
                    const t2Doc = await db.collection('trainers').doc(schoolData.trainerId2).get();
                    if (t2Doc.exists) trainer2 = { id: t2Doc.id, ...t2Doc.data() };
                }
                if (schoolData.coordinatorId) {
                    const coordDoc = await db.collection('coordinators').doc(schoolData.coordinatorId).get();
                    if (coordDoc.exists) coordinator = { id: coordDoc.id, ...coordDoc.data() };
                }
                schools.push({
                    id: doc.id,
                    ...schoolData,
                    eventDate,
                    trainer1Name: trainer1 ? trainer1.trainerName : 'N/A',
                    trainer2Name: trainer2 ? trainer2.trainerName : 'N/A',
                    coordinatorName: coordinator ? coordinator.name : 'N/A'
                });
            }
            console.log(`Fetched ${schools.length} schools`);

            const schoolColumns = [
                { header: 'Doc ID', key: 'id' },
                { header: 'School Name', key: 'schoolName' },
                { header: 'Registered At', key: 'registeredAt', transform: formatTimestamp },
                { header: 'Email', key: 'schoolEmail' },
                { header: 'City', key: 'city' },
                { header: 'District', key: 'district' },
                { header: 'Pincode', key: 'pincode' },
                { header: 'Event Date', key: 'eventDate', transform: formatTimestamp },
                { header: 'Approved', key: 'isApproved', transform: (val) => val ? 'Yes' : 'No' },
                { header: 'Completed', key: 'isCompleted', transform: (val) => val ? 'Yes' : 'No' },
                { header: 'Trainer 1', key: 'trainer1Name' },
                { header: 'Trainer 2', key: 'trainer2Name' },
                { header: 'Coordinator', key: 'coordinatorName' },
                { header: 'Principal Phone', key: 'principalNumber' },
                { header: 'Civics Teacher Phone', key: 'civicsTeacherNumber' },
                { header: 'School Phone', key: 'schoolPhoneNumber' },
                { header: 'Principal Email', key: 'principalEmail' },
                { header: 'Civics Teacher Email', key: 'civicsTeacherEmail' },
                { header: 'Resources Confirmed', key: 'resourcesConfirmed', transform: (val) => val ? 'Yes' : 'No' },
                { header: 'Selected Resources', key: 'selectedResources', transform: (val) => Array.isArray(val) ? val.join(', ') : 'None' },
                { header: 'Spot', key: 'spot' },
                { header: 'Reference Name', key: 'referenceName' }
            ];

            const schoolWorksheet = createWorksheet(schools, schoolColumns);
            XLSX.utils.book_append_sheet(workbook, schoolWorksheet, 'Schools');
        }

        if (type === 'trainers' || type === 'all') {
            // Fetch Trainers
            console.log('Fetching trainers...');
            trainersSnapshot = await db.collection('trainers').get();
            const trainers = trainersSnapshot.docs.map(doc => ({
                id: doc.id,
                ...doc.data(),
                registeredAt: doc.data().registeredAt?.toDate ? doc.data().registeredAt.toDate() : (doc.data().registeredAt ? new Date(doc.data().registeredAt) : null),
                isApproved: typeof doc.data().isApproved === 'boolean' ? doc.data().isApproved : false
            }));
            console.log(`Fetched ${trainers.length} trainers`);

          const trainerColumns = [
    { header: 'Doc ID', key: 'id' },
    { header: 'Name', key: 'trainerName' },
    { header: 'Email', key: 'email' },
    { header: 'City', key: 'city' },
    { header: 'District', key: 'district', transform: (val) => val || 'N/A' }, // Added district field
    { header: 'Profession', key: 'profession' },
    { header: 'Mobile Number', key: 'mobileNumber', transform: (val) => val || 'N/A' },
    { header: 'WhatsApp Number', key: 'whatsappNumber', transform: (val) => val || 'N/A' },
    { header: 'Reference Name', key: 'referenceName', transform: (val) => val || 'N/A' },
    { header: 'Approved', key: 'isApproved', transform: (val) => val ? 'Yes' : 'No' },
    { header: 'Registered At', key: 'registeredAt', transform: formatTimestamp }
];
            const trainerWorksheet = createWorksheet(trainers, trainerColumns);
            XLSX.utils.book_append_sheet(workbook, trainerWorksheet, 'Trainers');
        }

        if (type === 'students' || type === 'all') {
            // Fetch Participants
            console.log('Fetching participants...');
            const participantsSnapshot = await db.collection('participants').get();
            const participants = participantsSnapshot.docs.map(doc => {
                const data = doc.data();
                const score = Number(data.score) || null;
                const totalQuestions = Number(data.totalQuestions) || null;
                return {
                    id: doc.id,
                    studentName: data.studentName || 'N/A',
                    schoolNameDropdown: data.schoolNameDropdown || 'N/A',
                    studentClass: data.studentClass || 'N/A',
                    parentEmail: data.parentEmail || 'N/A',
                    parentMobile1: data.parentMobile1 || 'N/A',
                    parentMobile2: data.parentMobile2 || 'N/A',
                    address: data.address || 'N/A',
                    city: data.city || 'N/A',
                    pincode: data.pincode || 'N/A',
                    birthdate: data.birthdate?.toDate ? data.birthdate.toDate() : (data.birthdate ? new Date(data.birthdate) : null),
                    hasCompletedMCQ: data.hasCompletedMCQ || false,
                    score,
                    totalQuestions,
                    percentage: score && totalQuestions ? ((score / totalQuestions) * 100).toFixed(2) : null,
                    completedAt: data.completedAt?.toDate ? data.completedAt.toDate() : (data.completedAt ? new Date(data.completedAt) : null)
                };
            });
            console.log(`Fetched ${participants.length} participants`);

            const participantColumns = [
                { header: 'Doc ID', key: 'id' },
                { header: 'Name', key: 'studentName' },
                { header: 'School Name', key: 'schoolNameDropdown' },
                { header: 'Class', key: 'studentClass' },
                { header: 'Parent Email', key: 'parentEmail' },
                { header: 'Parent Mobile 1', key: 'parentMobile1' },
                { header: 'Parent Mobile 2', key: 'parentMobile2' },
                { header: 'Address', key: 'address' },
                { header: 'City', key: 'city' },
                { header: 'Pincode', key: 'pincode' },
                { header: 'Birthdate', key: 'birthdate', transform: formatTimestamp },
                { header: 'MCQ Completed', key: 'hasCompletedMCQ', transform: (val) => val ? 'Yes' : 'No' },
                { header: 'Score', key: 'score' },
                { header: 'Total Questions', key: 'totalQuestions' },
                { header: 'Percentage', key: 'percentage' },
                { header: 'Completed At', key: 'completedAt', transform: formatTimestamp }
            ];

            const participantWorksheet = createWorksheet(participants, participantColumns);
            XLSX.utils.book_append_sheet(workbook, participantWorksheet, 'Participants');
        }

        if (type === 'all') {
            // Fetch Coordinators
            console.log('Fetching coordinators...');
            const coordinatorsSnapshot = await db.collection('coordinators').get();
            const coordinators = coordinatorsSnapshot.docs.map(doc => ({
                id: doc.id,
                ...doc.data(),
                createdAt: doc.data().createdAt?.toDate() || null
            }));
            console.log(`Fetched ${coordinators.length} coordinators`);

            const coordinatorColumns = [
                { header: 'Doc ID', key: 'id' },
                { header: 'Name', key: 'name' },
                { header: 'Email', key: 'email' },
                { header: 'Number', key: 'number' },
                { header: 'Organization', key: 'organization' },
                { header: 'Created At', key: 'createdAt', transform: formatTimestamp }
            ];

            // Fetch Call Logs
            console.log('Fetching call logs...');
            const callLogsSnapshot = await db.collection('callLogs').get();
            const callLogs = callLogsSnapshot.docs.map(doc => ({
                id: doc.id,
                ...doc.data(),
                callDate: doc.data().callDate?.toDate ? doc.data().callDate.toDate() : (doc.data().callDate ? new Date(doc.data().callDate) : null)
            }));
            console.log(`Fetched ${callLogs.length} call logs`);

            const callLogColumns = [
                { header: 'Doc ID', key: 'id' },
                { header: 'Call Date', key: 'callDate', transform: formatTimestamp },
                { header: 'Caller', key: 'caller' },
                { header: 'Recipient', key: 'recipient' },
                { header: 'Duration', key: 'duration' },
                { header: 'Notes', key: 'notes' }
            ];

            // Fetch Media Uploads
            console.log('Fetching media uploads...');
            // Use schoolsSnapshot if it exists, otherwise fetch schools
            if (!schoolsSnapshot) {
                schoolsSnapshot = await db.collection('schools').get();
            }
            if (!trainersSnapshot) {
                trainersSnapshot = await db.collection('trainers').get();
            }

            const schoolMediaPromises = schoolsSnapshot.docs.map(doc =>
                db.collection('schools').doc(doc.id).collection('mediaUploads').get()
                    .then(mediaSnap => mediaSnap.docs.map(mediaDoc => ({
                        ...mediaDoc.data(),
                        uploadedAt: mediaDoc.data().uploadedAt?.toDate() || null,
                        uploadedBy: doc.data().schoolName || 'Unknown School'
                    })))
            );
            const trainerMediaPromises = trainersSnapshot.docs.map(doc =>
                db.collection('trainers').doc(doc.id).collection('mediaUploads').get()
                    .then(mediaSnap => mediaSnap.docs.map(mediaDoc => ({
                        ...mediaDoc.data(),
                        uploadedAt: mediaDoc.data().uploadedAt?.toDate() || null,
                        uploadedBy: doc.data().trainerName || 'Unknown Trainer'
                    })))
            );
            const mediaUploads = [...(await Promise.all(schoolMediaPromises)).flat(), ...(await Promise.all(trainerMediaPromises)).flat()];
            console.log(`Fetched ${mediaUploads.length} media uploads`);

            const mediaColumns = [
                { header: 'Media ID', key: 'mediaId' },
                { header: 'Uploaded By', key: 'uploadedBy' },
                { header: 'Uploaded At', key: 'uploadedAt', transform: formatTimestamp },
                { header: 'Description', key: 'description' },
                { header: 'Type', key: 'type' },
                { header: 'Seen', key: 'seen', transform: (val) => val ? 'Yes' : 'No' }
            ];

            const coordinatorWorksheet = createWorksheet(coordinators, coordinatorColumns);
            const callLogWorksheet = createWorksheet(callLogs, callLogColumns);
            const mediaWorksheet = createWorksheet(mediaUploads, mediaColumns);

            XLSX.utils.book_append_sheet(workbook, coordinatorWorksheet, 'Coordinators');
            XLSX.utils.book_append_sheet(workbook, callLogWorksheet, 'Call Logs');
            XLSX.utils.book_append_sheet(workbook, mediaWorksheet, 'Media Uploads');
        }

        console.log('Generating buffer...');
        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

        console.log('Sending response...');
        res.setHeader('Content-Disposition', `attachment; filename=${type === 'all' ? 'admin_dashboard' : type}_data_${new Date().toISOString().split('T')[0]}.xlsx`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (error) {
        console.error('Error generating Excel file:', error.message, error.stack);
        res.status(500).json({ error: `Failed to generate Excel file: ${error.message}` });
    }
});
    
// Admin dashboard (renders adminDashboard.ejs)

app.get('/admin-dashboard', requireAdmin, async (req, res) => {
    try {
        console.log('---- Admin Dashboard Route Triggered ----');

        // Helper to format Firestore Timestamp safely
        const formatTime = (timestamp) => {
            if (!timestamp) return null;
            try {
                const dateObj = timestamp.toDate ? timestamp.toDate() : new Date(timestamp);
                if (isNaN(dateObj.getTime())) return null;
                return dateObj.toLocaleTimeString('en-IN', {
                    hour: '2-digit',
                    minute: '2-digit',
                    hour12: true
                });
            } catch (err) {
                console.warn('Error formatting timestamp:', err.message);
                return null;
            }
        };

        // Query parameter filters
        const filters = {
            schoolApproved: req.query.schoolApproved === 'true' ? true : req.query.schoolApproved === 'false' ? false : '',
            schoolCity: req.query.schoolCity?.trim() || '',
            trainerApproved: req.query.trainerApproved === 'true' ? true : req.query.trainerApproved === 'false' ? false : '',
            trainerCity: req.query.trainerCity?.trim() || '',
            studentCompleted: req.query.studentCompleted === 'true' ? true : req.query.studentCompleted === 'false' ? false : '',
            studentCity: req.query.studentCity?.trim() || '',
            workshopCompleted: req.query.workshopCompleted === 'true' ? true : req.query.workshopCompleted === 'false' ? false : '',
            coordinatorId: req.query.coordinatorId?.trim() || ''
        };

        /* ------------------- FETCH SCHOOLS ------------------- */
        let schoolsQuery = db.collection('schools');
        if (filters.schoolApproved !== '') schoolsQuery = schoolsQuery.where('isApproved', '==', filters.schoolApproved);
        if (filters.schoolCity) schoolsQuery = schoolsQuery.where('city', '==', filters.schoolCity);
        if (filters.workshopCompleted !== '') schoolsQuery = schoolsQuery.where('workshopCompleted', '==', filters.workshopCompleted);

        const schoolsSnapshot = await schoolsQuery.get();
        if (schoolsSnapshot.empty) console.warn('No schools found in Firestore');

        const schools = [];
        const schoolMediaPromises = [];
        const trainerIds = new Set();
        const coordinatorIds = new Set();   

        for (const doc of schoolsSnapshot.docs) {
            const schoolData = doc.data();
            let eventDate = schoolData.eventDate
                ? schoolData.eventDate.toDate
                    ? schoolData.eventDate.toDate()
                    : new Date(schoolData.eventDate)
                : null;
            if (eventDate && isNaN(eventDate.getTime())) eventDate = null;

            const workshopStartTime = formatTime(schoolData.workshopStartTime);
            const workshopEndTime = formatTime(schoolData.workshopEndTime);

            if (schoolData.trainerId1) trainerIds.add(schoolData.trainerId1);
            if (schoolData.trainerId2) trainerIds.add(schoolData.trainerId2);
            if (schoolData.coordinatorId) coordinatorIds.add(schoolData.coordinatorId);

            schoolMediaPromises.push(
                db.collection('schools').doc(doc.id).collection('mediaUploads')
                    .where('type', 'in', ['image', 'video'])
                    .get()
                    .then(mediaSnap => mediaSnap.docs.map(mediaDoc => ({
                        id: mediaDoc.id,
                        ...mediaDoc.data(),
                        uploadedAt: mediaDoc.data().uploadedAt?.toDate() || null,
                        uploadedBy: schoolData.schoolName || 'Unknown School',
                        uploaderType: 'School',
                        source: 'school'
                    })))
                    .catch(err => {
                        console.error(`Error fetching media for school ${doc.id}:`, err.message);
                        return [];
                    })
            );

            schools.push({
                id: doc.id,
                ...schoolData,
                eventDate,
                workshopStartTime,
                workshopEndTime
            });
        }

        const trainerDocs = await Promise.all(
            Array.from(trainerIds).map(id => db.collection('trainers').doc(id).get())
        );
        const trainersMap = new Map(trainerDocs.map(doc => [doc.id, doc.exists ? { id: doc.id, ...doc.data() } : null]));

        const coordinatorDocs = await Promise.all(
            Array.from(coordinatorIds).map(id => db.collection('coordinators').doc(id).get())
        );
        const coordinatorsMap = new Map(coordinatorDocs.map(doc => [doc.id, doc.exists ? { id: doc.id, ...doc.data() } : null]));

        schools.forEach(school => {
            school.trainer1 = school.trainerId1 ? trainersMap.get(school.trainerId1) : null;
            school.trainer2 = school.trainerId2 ? trainersMap.get(school.trainerId2) : null;
            school.coordinator = school.coordinatorId ? coordinatorsMap.get(school.coordinatorId) : null;
        });

        /* ------------------- FETCH TRAINERS ------------------- */
        let trainersQuery = db.collection('trainers');
        if (filters.trainerApproved !== '') trainersQuery = trainersQuery.where('isApproved', '==', filters.trainerApproved);
        if (filters.trainerCity) trainersQuery = trainersQuery.where('city', '==', filters.trainerCity);

        const trainersSnapshot = await trainersQuery.get();
        if (trainersSnapshot.empty) console.warn('No trainers found in Firestore');

        const trainers = [];
        const trainerMediaPromises = [];

        for (const doc of trainersSnapshot.docs) {
            const data = doc.data();
            trainerMediaPromises.push(
                db.collection('trainers').doc(doc.id).collection('mediaUploads')
                    .where('type', 'in', ['image', 'video'])
                    .get()
                    .then(mediaSnap => mediaSnap.docs.map(mediaDoc => ({
                        id: mediaDoc.id,
                        ...mediaDoc.data(),
                        uploadedAt: mediaDoc.data().uploadedAt?.toDate() || null,
                        uploadedBy: data.trainerName || 'Unknown Trainer',
                        uploaderType: 'Trainer',
                        source: 'trainer'
                    })))
                    .catch(err => {
                        console.error(`Error fetching media for trainer ${doc.id}:`, err.message);
                        return [];
                    })
            );

            let registeredAt = data.registeredAt
                ? data.registeredAt.toDate
                    ? data.registeredAt.toDate()
                    : new Date(data.registeredAt)
                : null;
            if (registeredAt && isNaN(registeredAt.getTime())) registeredAt = null;

            trainers.push({
                id: doc.id,
                ...data,
                registeredAt,
                isApproved: typeof data.isApproved === 'boolean' ? data.isApproved : false
            });
        }

        /* ------------------- FETCH COORDINATORS ------------------- */
        const coordinatorsSnapshot = await db.collection('coordinators').get();
        if (coordinatorsSnapshot.empty) console.warn('No coordinators found in Firestore');

        const coordinators = coordinatorsSnapshot.docs.map(doc => {
            const data = doc.data();
            let createdAt = data.createdAt
                ? data.createdAt.toDate
                    ? data.createdAt.toDate()
                    : new Date(data.createdAt)
                : null;
            if (createdAt && isNaN(createdAt.getTime())) createdAt = null;
            return {
                id: doc.id,
                ...data,
                createdAt
            };
        });

        /* ------------------- FETCH PARTICIPANTS ------------------- */
        let participantsQuery = db.collection('participants');
        if (filters.studentCompleted !== '') participantsQuery = participantsQuery.where('hasCompletedMCQ', '==', filters.studentCompleted);
        if (filters.studentCity) participantsQuery = participantsQuery.where('city', '==', filters.studentCity);

        const participantsSnapshot = await participantsQuery.get();
        if (participantsSnapshot.empty) console.warn('No participants found in Firestore');

        const participants = [];
        const studentMediaPromises = [];

        for (const doc of participantsSnapshot.docs) {
            const data = doc.data();
            studentMediaPromises.push(
                db.collection('participants').doc(doc.id).collection('mediaUploads')
                    .where('type', 'in', ['image', 'video'])
                    .get()
                    .then(mediaSnap => mediaSnap.docs.map(mediaDoc => ({
                        id: mediaDoc.id,
                        ...mediaDoc.data(),
                        uploadedAt: mediaDoc.data().uploadedAt?.toDate() || null,
                        uploadedBy: data.studentName || 'Unknown Student',
                        uploaderType: 'Student',
                        source: 'student'
                    })))
                    .catch(err => {
                        console.error(`Error fetching media for participant ${doc.id}:`, err.message);
                        return [];
                    })
            );

            const score = Number(data.score) || null;
            const totalQuestions = Number(data.totalQuestions) || null;
            const trial1Score = Number(data.trial1Score) || null;
            const trial1TotalQuestions = Number(data.trial1TotalQuestions) || null;
            const trial2Score = Number(data.trial2Score) || null;
            const trial2TotalQuestions = Number(data.trial2TotalQuestions) || null;

            let birthdate = data.birthdate ? (data.birthdate.toDate ? data.birthdate.toDate() : new Date(data.birthdate)) : null;
            if (birthdate && isNaN(birthdate.getTime())) birthdate = null;

            let completedAt = data.completedAt ? (data.completedAt.toDate ? data.completedAt.toDate() : new Date(data.completedAt)) : null;
            if (completedAt && isNaN(completedAt.getTime())) completedAt = null;

            participants.push({
                id: doc.id,
                ...data,
                birthdate,
                completedAt,
                score,
                totalQuestions,
                percentage: score && totalQuestions ? ((score / totalQuestions) * 100).toFixed(2) : null,
                trial1Score,
                trial1TotalQuestions,
                trial1Percentage: trial1Score && trial1TotalQuestions ? ((trial1Score / trial1TotalQuestions) * 100).toFixed(2) : null,
                trial2Score,
                trial2TotalQuestions,
                trial2Percentage: trial2Score && trial2TotalQuestions ? ((trial2Score / trial2TotalQuestions) * 100).toFixed(2) : null,
                trial1CorrectAnswers: Number(data.trial1CorrectAnswers) || null,
                trial1WrongAnswers: Number(data.trial1WrongAnswers) || null,
                trial2CorrectAnswers: Number(data.trial2CorrectAnswers) || null,
                trial2WrongAnswers: Number(data.trial2WrongAnswers) || null
            });
        }

        /* ------------------- FETCH CALL LOGS ------------------- */
        const callLogsSnapshot = await db.collection('callLogs').orderBy('createdAt', 'desc').get();
        if (callLogsSnapshot.empty) console.warn('No call logs found in Firestore');

        const callLogs = callLogsSnapshot.docs.map(doc => {
            const data = doc.data();
            let callDate = data.callDate ? (data.callDate.toDate ? data.callDate.toDate() : new Date(data.callDate)) : null;
            if (callDate && isNaN(callDate.getTime())) callDate = null;
            let createdAt = data.createdAt ? (data.createdAt.toDate ? data.createdAt.toDate() : new Date(data.createdAt)) : null;
            if (createdAt && isNaN(createdAt.getTime())) createdAt = null;
            return { id: doc.id, ...data, callDate, createdAt };
        });

        /* ------------------- FETCH WORKSHOP SUMMARIES ------------------- */
        let workshopSummariesQuery = db.collection('workshopSummaries');
        if (filters.coordinatorId) workshopSummariesQuery = workshopSummariesQuery.where('coordinatorId', '==', filters.coordinatorId);

        const workshopSummariesSnapshot = await workshopSummariesQuery.get();
        if (workshopSummariesSnapshot.empty) console.warn('No workshop summaries found in Firestore');

        const workshopSummaries = workshopSummariesSnapshot.docs.map(doc => {
            const data = doc.data();
            let createdAt = data.createdAt
                ? data.createdAt.toDate
                    ? data.createdAt.toDate()
                    : new Date(data.createdAt)
                : null;
            if (createdAt && isNaN(createdAt.getTime())) createdAt = null;

            return {
                id: doc.id,
                schoolName: data.schoolName || '',
                trainer1: data.trainer1 || '',
                trainer2: data.trainer2 || '',
                coordinatorName: data.coordinatorName || '',
                workshopDate: data.workshopDate || '',
                financialStatus: data.financialStatus || '',
                trainerFeedback: data.trainerFeedback || '', // Include trainerFeedback
                paymentMode: data.paymentMode || '',
                upiId: data.upiId || '',
                transactionId: data.transactionId || '',
                createdAt: createdAt ? createdAt.toLocaleDateString('en-IN') : 'N/A'
            };
        });

        /* ------------------- FETCH VISITED SCHOOLS ------------------- */
        const visitedSchoolsSnapshot = await db.collection('visitedSchools').get();
        if (visitedSchoolsSnapshot.empty) console.warn('No visited schools found in Firestore');

        const visitedSchools = visitedSchoolsSnapshot.docs.map(doc => {
            const data = doc.data();
            let visitDate = data.visitDate ? (data.visitDate.toDate ? data.visitDate.toDate() : new Date(data.visitDate)) : null;
            let createdAt = data.createdAt ? (data.createdAt.toDate ? data.createdAt.toDate() : new Date(data.createdAt)) : null;
            return {
                id: doc.id,
                name: data.name || data.schoolName || 'N/A',
                visitDate: visitDate ? visitDate.toLocaleDateString('en-IN') : 'N/A',
                schoolAddress: data.schoolAddress || 'N/A',
                contactPerson: data.contactPerson || 'N/A',
                contactNumber: data.contactNumber || 'N/A',
                visitNotes: data.visitNotes || data.remarks || 'N/A',
                coordinatorId: data.coordinatorId || 'N/A',
                createdAt: createdAt ? createdAt.toLocaleDateString('en-IN') : 'N/A'
            };
        });

        // Resolve coordinator names for visited schools
        const vsCoordinatorIds = new Set(visitedSchools.map(v => v.coordinatorId).filter(id => id && id !== 'N/A'));
        const vsCoordinatorDocs = await Promise.all(Array.from(vsCoordinatorIds).map(id => db.collection('coordinators').doc(id).get()));
        const vsCoordinatorsMap = new Map(vsCoordinatorDocs.filter(doc => doc.exists).map(doc => [doc.id, doc.data().name || doc.data().coordinatorName || 'Unknown']));
        visitedSchools.forEach(vs => { vs.coordinatorName = vsCoordinatorsMap.get(vs.coordinatorId) || 'N/A'; });

        /* ------------------- FETCH STUDENT FEEDBACK ------------------- */
        let feedbackEntries = [];
        try {
            let feedbackQuery = db.collection('studentFeedback');
            if (filters.coordinatorId) feedbackQuery = feedbackQuery.where('coordinatorId', '==', filters.coordinatorId);
            const feedbackSnapshot = await feedbackQuery.where('status', '==', 'submitted').orderBy('createdAt', 'desc').get();
            feedbackEntries = feedbackSnapshot.docs.map(doc => {
                const data = doc.data();
                return {
                    id: doc.id,
                    studentName: data.studentName || 'N/A',
                    className: data.className || 'N/A',
                    question1: data.question1 || 'N/A',
                    question2: data.question2 || 'N/A',
                    question3: data.question3 || 'N/A',
                    question4: data.question4 || 'N/A',
                    question5: data.question5 || 'N/A',
                    coordinatorId: data.coordinatorId || 'N/A',
                    status: data.status || 'N/A'
                };
            });
        } catch (feedbackErr) {
            console.error('Error fetching student feedback:', feedbackErr);
        }

        /* ------------------- COMBINE MEDIA UPLOADS ------------------- */
        const schoolMedia = (await Promise.all(schoolMediaPromises)).flat();
        const trainerMedia = (await Promise.all(trainerMediaPromises)).flat();
        const studentMedia = (await Promise.all(studentMediaPromises)).flat();
        const imageUploads = {
            images: {
                school: schoolMedia.filter(m => m.type === 'image'),
                trainer: trainerMedia.filter(m => m.type === 'image'),
                student: studentMedia.filter(m => m.type === 'image')
            },
            videos: {
                school: schoolMedia.filter(m => m.type === 'video'),
                trainer: trainerMedia.filter(m => m.type === 'video'),
                student: studentMedia.filter(m => m.type === 'video')
            },
            all: [...schoolMedia, ...trainerMedia, ...studentMedia]
                .filter(m => m.type === 'image' || m.type === 'video')
                .sort((a, b) => (b.uploadedAt || 0) - (a.uploadedAt || 0))
        };

        /* ------------------- TODAY STATS ------------------- */
        const today = new Date(); today.setHours(0, 0, 0, 0);
        const tomorrow = new Date(today); tomorrow.setDate(today.getDate() + 1);
        let studentsTodayCount = 0, averageScore = 0;

        try {
            const studentsTodaySnapshot = await db.collection('participants')
                .where('hasCompletedMCQ', '==', true)
                .where('completedAt', '>=', today)
                .where('completedAt', '<', tomorrow)
                .get();
            studentsTodayCount = studentsTodaySnapshot.size;
            const totalScore = studentsTodaySnapshot.docs.reduce((sum, doc) => sum + (Number(doc.data().score) || 0), 0);
            averageScore = studentsTodayCount > 0 ? (totalScore / studentsTodayCount).toFixed(2) : 0;
        } catch (indexErr) {
            console.error('🔥 Firestore Index Error in Today Stats:', indexErr);
        }
            /* ------------------- pdf get ------------------- */
const pdfUploadsSnapshot = await db.collection('pdfUploads').get();

const pdfs = pdfUploadsSnapshot.docs.map(doc => {
    const data = doc.data();
    return {
        id: doc.id,
        name: data.name || 'Unnamed PDF',
        url: data.url || '',
        uploadedAt: data.uploadedAt?.toDate ? data.uploadedAt.toDate() : null
    };
});
        /* ------------------- RENDER PAGE ------------------- */
        res.render('adminDashboard', {
            schools,
            trainers,
            coordinators,
            participants,
            callLogs,
            filters,
            studentsTodayCount,
            averageScore,
            mediaUploads: imageUploads.all,
            imageUploads,
            workshopSummaries,
            feedbackEntries,
            visitedSchools,
             pdfs,
            error: null,
            success: req.query.success || null,
            showRegister: false
        });

    } catch (err) {
        console.error('❌ Error in /admin-dashboard:', err);
        res.render('adminDashboard', {
            schools: [],
            trainers: [],
            coordinators: [],
            participants: [],
            callLogs: [],
            filters: {},
            studentsTodayCount: 0,
            averageScore: 0,
            mediaUploads: [],
            imageUploads: { images: { school: [], trainer: [], student: [] }, videos: { school: [], trainer: [], student: [] }, all: [] },
            workshopSummaries: [],
            feedbackEntries: [],
            visitedSchools: [],
            error: `Failed to load dashboard data: ${err.message}`,
            success: null,
            showRegister: false
        });
    }
});

//update-school-status
app.post('/update-school-status', requireAdmin, async (req, res) => {
    try {
        const { schoolId, thinkingLetter, cotation } = req.body;

        // Validate schoolId
        if (!schoolId) {
            return res.status(400).json({ error: 'Missing schoolId' });
        }

        const updateData = {};

        // Validate and process thinkingLetter
        if (thinkingLetter !== undefined) {
            if (!['Yes', 'No'].includes(thinkingLetter)) {
                return res.status(400).json({ error: 'thinkingLetter must be "Yes" or "No"' });
            }
            updateData.thinkingLetter = thinkingLetter === 'Yes';
        }

        // Validate and process cotation
        if (cotation !== undefined) {
            if (!['Yes', 'No'].includes(cotation)) {
                return res.status(400).json({ error: 'cotation must be "Yes" or "No"' });
            }
            updateData.cotation = cotation === 'Yes';
        }

        // Check if there are fields to update
        if (Object.keys(updateData).length === 0) {
            return res.status(400).json({ error: 'No valid fields to update' });
        }

        // Update the Firestore document
        await db.collection('schools').doc(schoolId).update(updateData);

        res.status(200).json({ success: true });
    } catch (err) {
        console.error('Error updating school status:', err);
        res.status(500).json({ error: 'Failed to update status' });
    }
});
///admin/assign-coordinator/:schoolId
app.post('/admin/assign-coordinator/:schoolId', requireAdmin, async (req, res) => {
    try {
        const schoolId = req.params.schoolId;
        const { coordinatorId } = req.body;

        if (!coordinatorId) {
            return res.status(400).json({ message: 'Coordinator ID is required' });
        }

        // Verify coordinator exists
        const coordinatorDoc = await db.collection('coordinators').doc(coordinatorId).get();
        if (!coordinatorDoc.exists) {
            return res.status(404).json({ message: 'Coordinator not found' });
        }

        // Update school with coordinatorId
        await db.collection('schools').doc(schoolId).update({
            coordinatorId,
            updatedAt: admin.firestore.FieldValue.serverTimestamp()
        });

        console.log(`Coordinator ${coordinatorId} assigned to school ${schoolId}`);
        res.redirect('/admin-dashboard?success=Coordinator assigned successfully');
    } catch (error) {
        console.error('Error assigning coordinator:', error.message, error.stack);
        res.redirect('/admin-dashboard?error=Failed to assign coordinator');
    }
});
    // Delete a participant by ID
    app.post('/admin-dashboard/delete-student/:id', requireAdmin, async (req, res) => {
        const participantId = req.params.id;

        try {
            await db.collection('participants').doc(participantId).delete();
            res.redirect('/admin-dashboard?success=Student deleted successfully');
        } catch (err) {
            console.error('❌ Error deleting student:', err.message);
            res.redirect('/admin-dashboard?error=Failed to delete student');
        }
    });
    // ✅ POST mark media as seen (One-way, not toggle)
    app.post('/admin-dashboard/mark-media-seen/:mediaId', async (req, res) => {
        try {
            const mediaId = req.params.mediaId;
            const mediaSnap = await db.collectionGroup('mediaUploads')
                .where('id', '==', mediaId)
                .limit(1)
                .get();

            if (mediaSnap.empty) {
                return res.status(404).json({ error: 'Media not found' });
            }

            const mediaDoc = mediaSnap.docs[0];

            if (!mediaDoc.data().seen) {
                await mediaDoc.ref.update({ seen: true });
            }

            res.json({ seen: true, message: 'Media marked as seen' });
        } catch (err) {
            console.error('Error marking media as seen:', err.message);
            res.status(500).json({ error: 'Server error' });
        }
    });

    app.get('/admin-dashboard/get-media-status/:mediaId', async (req, res) => {
        try {
            const mediaId = req.params.mediaId;
            const mediaSnap = await db.collectionGroup('mediaUploads').where('id', '==', mediaId).limit(1).get();

            if (mediaSnap.empty) {
                return res.status(404).json({ error: 'Media not found' });
            }

            const mediaDoc = mediaSnap.docs[0];
            res.json({ seen: mediaDoc.data().seen || false });
        } catch (err) {
            console.error('Error fetching media status:', err.message);
            res.status(500).json({ error: 'Server error' });
        }
    });

// POST: Toggle media seen status
app.post('/admin-dashboard/toggle-media-seen/:mediaId', async (req, res) => {
    try {
        const mediaId = req.params.mediaId;
        const mediaSnap = await db.collectionGroup('mediaUploads').where('id', '==', mediaId).limit(1).get();

        if (mediaSnap.empty) {
            return res.status(404).json({ error: 'Media not found' });
        }

        const mediaDoc = mediaSnap.docs[0];
        const currentSeen = mediaDoc.data().seen || false;
        const newSeen = !currentSeen;

        await mediaDoc.ref.update({ seen: newSeen });

        res.json({ seen: newSeen });
    } catch (err) {
        console.error('Error toggling media status:', err.message);
        res.status(500).json({ error: 'Server error' });
    }
});
app.post('/admin-dashboard/delete-trainer/:id', async (req, res) => {
  const trainerId = req.params.id;

  try {
    await db.collection('trainers').doc(trainerId).delete();
    res.redirect('/admin-dashboard?success=Trainer deleted successfully');
  } catch (err) {
    console.error('Error deleting trainer:', err.message);
    res.redirect('/admin-dashboard?error=Failed to delete trainer');
  }
});

    // Approve trainer (general approval, not tied to a specific school)
    app.post('/admin/approve-trainer', requireAdmin, async (req, res) => {
        try {
            const { trainerId } = req.body;
            if (!trainerId) {
                return res.redirect('/admin-dashboard?error=Trainer ID is required');
            }

            const trainerRef = db.collection('trainers').doc(trainerId);
            const trainerDoc = await trainerRef.get();
            if (!trainerDoc.exists) {
                return res.redirect('/admin-dashboard?error=Trainer not found');
            }

            await trainerRef.update({ isApproved: true });
            res.redirect('/admin-dashboard?success=Trainer approved successfully');
        } catch (error) {
            console.error('Error in admin/approve-trainer route:', error.message, error.stack);
            res.redirect('/admin-dashboard?error=Error approving trainer');
        }
    });

    // Admin approve school
    app.post('/admin/approve-school/:id', requireAdmin, async (req, res) => {
            try {
                const schoolId = req.params.id;
                await db.collection('schools').doc(schoolId).update({ isApproved: true });
                res.redirect('/admin-dashboard');
            } catch (error) {
                console.error('Error in approve-school route:', error.message, error.stack);
                res.status(500).send('Error approving school.');
            }
        });

    // Admin reject school
    app.post('/admin/reject-school', requireAdmin, async (req, res) => {
    try {
    const { schoolId } = req.body;
    if (!schoolId) {
    return res.redirect('/admin-dashboard?error=School ID is required');
    }

    const schoolRef = db.collection('schools').doc(schoolId);
    const schoolDoc = await schoolRef.get();
    if (!schoolDoc.exists) {
    return res.redirect('/admin-dashboard?error=School not found');
    }

    await schoolRef.delete();
    res.redirect('/admin-dashboard?success=School rejected successfully');
    } catch (error) {
    console.error('Error in admin/reject-school route:', error.message, error.stack);
    res.redirect('/admin-dashboard?error=Error rejecting school');
    }
    });

    // Admin approve trainer
    // Approve Trainer 1 for a specific school
    app.post('/admin/approve-trainer1/:schoolId', requireAdmin, async (req, res) => {
        try {
            const { schoolId } = req.params;
            const schoolRef = db.collection('schools').doc(schoolId);
            const schoolDoc = await schoolRef.get();
            if (!schoolDoc.exists) {
                return res.redirect('/admin-dashboard?error=School not found');
            }

            const { trainerId1, name: schoolName, eventDate, spot, facilities } = schoolDoc.data();
            if (!trainerId1) {
                return res.redirect('/admin-dashboard?error=No Trainer 1 assigned to approve');
            }

            const trainerRef = db.collection('trainers').doc(trainerId1);
            const trainerDoc = await trainerRef.get();
            if (!trainerDoc.exists) {
                return res.redirect('/admin-dashboard?error=Trainer 1 not found');
            }

            // Check if trainer is already assigned to another school
            const existingSchoolId = trainerDoc.data().assignedSchoolId;
            if (existingSchoolId && existingSchoolId !== schoolId) {
                return res.redirect('/admin-dashboard?error=Trainer 1 already assigned to another school');
            }

            // Update trainer with approval status and assigned school details
            await trainerRef.update({
                isApproved: true,
                assignedSchoolId: schoolId,
                assignedSchoolName: schoolName
            });

            // Create assignment subcollection document
            await trainerRef.collection('assignments').doc(schoolId).set({
                schoolId,
                schoolName,
                eventDate: eventDate || null,
                spot: spot || null,
                facilities: facilities || [],
                assignedAt: admin.firestore.Timestamp.now()
            });

            console.log(`Approved Trainer 1 ${trainerId1} for school ${schoolId} (${schoolName})`);
            res.redirect('/admin-dashboard?success=Trainer 1 approved successfully');
        } catch (error) {
            console.error('Error approving Trainer 1:', error.message, error.stack);
            res.redirect('/admin-dashboard?error=Failed to approve Trainer 1');
        }
    });

    // Approve Trainer 2 for a specific school
    app.post('/admin/approve-trainer2/:schoolId', requireAdmin, async (req, res) => {
        try {
            const { schoolId } = req.params;
            const schoolRef = db.collection('schools').doc(schoolId);
            const schoolDoc = await schoolRef.get();
            if (!schoolDoc.exists) {
                return res.redirect('/admin-dashboard?error=School not found');
            }

            const { trainerId2, name: schoolName, eventDate, spot, facilities } = schoolDoc.data();
            if (!trainerId2) {
                return res.redirect('/admin-dashboard?error=No Trainer 2 assigned to approve');
            }

            const trainerRef = db.collection('trainers').doc(trainerId2);
            const trainerDoc = await trainerRef.get();
            if (!trainerDoc.exists) {
                return res.redirect('/admin-dashboard?error=Trainer 2 not found');
            }

            // Check if trainer is.cn already assigned to another school
            const existingSchoolId = trainerDoc.data().assignedSchoolId;
            if (existingSchoolId && existingSchoolId !== schoolId) {
                return res.redirect('/admin-dashboard?error=Trainer 2 already assigned to another school');
            }

            // Update trainer with approval status and assigned school details
            await trainerRef.update({
                isApproved: true,
                assignedSchoolId: schoolId,
                assignedSchoolName: schoolName
            });

            // Create assignment subcollection document
            await trainerRef.collection('assignments').doc(schoolId).set({
                schoolId,
                schoolName,
                eventDate: eventDate || null,
                spot: spot || null,
                facilities: facilities || [],
                assignedAt: admin.firestore.Timestamp.now()
            });

            console.log(`Approved Trainer 2 ${trainerId2} for school ${schoolId} (${schoolName})`);
            res.redirect('/admin-dashboard?success=Trainer 2 approved successfully');
        } catch (error) {
            console.error('Error approving Trainer 2:', error.message, error.stack);
            res.redirect('/admin-dashboard?error=Failed to approve Trainer 2');
        }
    });

    // Reject Trainer 1
    app.post('/admin/reject-trainer1/:schoolId', requireAdmin, async (req, res) => {
        try {
            const { schoolId } = req.params;
            const schoolRef = db.collection('schools').doc(schoolId);
            await schoolRef.update({
                trainerId1: null
            });
            console.log(`Rejected Trainer 1 for school ${schoolId}`);
            res.redirect('/admin-dashboard?success=Trainer 1 rejected successfully');
        } catch (error) {
            console.error('Error rejecting Trainer 1:', error);
            res.redirect('/admin-dashboard?error=Failed to reject Trainer 1');
        }
    });

    // Reject Trainer 2
    app.post('/admin/reject-trainer2/:schoolId', requireAdmin, async (req, res) => {
        try {
            const { schoolId } = req.params;
            const schoolRef = db.collection('schools').doc(schoolId);
            await schoolRef.update({
                trainerId2: null
            });
            console.log(`Rejected Trainer 2 for school ${schoolId}`);
            res.redirect('/admin-dashboard?success=Trainer 2 rejected successfully');
        } catch (error) {
            console.error('Error rejecting Trainer 2:', error);
            res.redirect('/admin-dashboard?error=Failed to reject Trainer 2');
        }
    });

    // Assign Trainer 1
    app.post('/admin/assign-trainer1/:schoolId', requireAdmin, async (req, res) => {
        try {
            const { schoolId } = req.params;
            const { trainerId } = req.body;

            if (!trainerId) {
                return res.redirect('/admin-dashboard?error=Trainer ID is required for Trainer 1');
            }

            const schoolRef = db.collection('schools').doc(schoolId);
            const schoolDoc = await schoolRef.get();
            if (!schoolDoc.exists) {
                return res.redirect('/admin-dashboard?error=School not found');
            }

            const { name: schoolName, eventDate, spot, facilities } = schoolDoc.data();
            const trainerRef = db.collection('trainers').doc(trainerId);
            const trainerDoc = await trainerRef.get();
            if (!trainerDoc.exists) {
                return res.redirect('/admin-dashboard?error=Trainer not found');
            }

            // Check if trainer is already assigned to another school
            const existingSchoolId = trainerDoc.data().assignedSchoolId;
            if (existingSchoolId && existingSchoolId !== schoolId) {
                return res.redirect('/admin-dashboard?error=Trainer already assigned to another school');
            }

            // Update school with trainer ID
            await schoolRef.update({
                trainerId1: trainerId
            });

            // Update trainer with assigned school details
            await trainerRef.update({
                assignedSchoolId: schoolId,
                assignedSchoolName: schoolName
            });

            // Create assignment subcollection document
            await trainerRef.collection('assignments').doc(schoolId).set({
                schoolId,
                schoolName,
                eventDate: eventDate || null,
                spot: spot || null,
                facilities: facilities || [],
                assignedAt: admin.firestore.Timestamp.now()
            });

            console.log(`Assigned Trainer 1 ${trainerId} to school ${schoolId}`);
            res.redirect('/admin-dashboard?success=Trainer 1 assigned successfully');
        } catch (error) {
            console.error('Error assigning Trainer 1:', error);
            res.redirect('/admin-dashboard?error=Failed to assign Trainer 1');
        }
    });

    // Assign Trainer 2
    app.post('/admin/assign-trainer2/:schoolId', requireAdmin, async (req, res) => {
        try {
            const { schoolId } = req.params;
            const { trainerId } = req.body;

            if (!trainerId) {
                return res.redirect('/admin-dashboard?error=Trainer ID is required for Trainer 2');
            }

            const schoolRef = db.collection('schools').doc(schoolId);
            const schoolDoc = await schoolRef.get();
            if (!schoolDoc.exists) {
                return res.redirect('/admin-dashboard?error=School not found');
            }

            const { name: schoolName, eventDate, spot, facilities } = schoolDoc.data();
            const trainerRef = db.collection('trainers').doc(trainerId);
            const trainerDoc = await trainerRef.get();
            if (!trainerDoc.exists) {
                return res.redirect('/admin-dashboard?error=Trainer not found');
            }

            // Check if trainer is already assigned to another school
            const existingSchoolId = trainerDoc.data().assignedSchoolId;
            if (existingSchoolId && existingSchoolId !== schoolId) {
                return res.redirect('/admin-dashboard?error=Trainer already assigned to another school');
            }

            // Update school with trainer ID
            await schoolRef.update({
                trainerId2: trainerId
            });

            // Update trainer with assigned school details
            await trainerRef.update({
                assignedSchoolId: schoolId,
                assignedSchoolName: schoolName
            });

            // Create assignment subcollection document
            await trainerRef.collection('assignments').doc(schoolId).set({
                schoolId,
                schoolName,
                eventDate: eventDate || null,
                spot: spot || null,
                facilities: facilities || [],
                assignedAt: admin.firestore.Timestamp.now()
            });

            console.log(`Assigned Trainer 2 ${trainerId} to school ${schoolId}`);
            res.redirect('/admin-dashboard?success=Trainer 2 assigned successfully');
        } catch (error) {
            console.error('Error assigning Trainer 2:', error);
            res.redirect('/admin-dashboard?error=Failed to assign Trainer 2');
        }
    });

    // Reject trainer and show available trainers (for assigning new trainers)
    app.post('/reject-trainer/:id', requireAdmin, async (req, res) => {
        try {
            const availableTrainersSnapshot = await db.collection('trainers')
                .where('isApproved', '==', true)
                .get();
            const availableTrainers = availableTrainersSnapshot.docs.map(doc => ({
                id: doc.id,
                ...doc.data()
            }));
            res.render('assignTrainer', {
                rejectedTrainerId: req.params.id,
                availableTrainers,
                error: null,
                schoolId: req.query.schoolId || null
            });
        } catch (error) {
            console.error('Error fetching available trainers:', error);
            res.render('assignTrainer', {
                rejectedTrainerId: req.params.id,
                availableTrainers: [],
                error: 'Failed to load available trainers',
                schoolId: req.query.schoolId || null
            });
        }
    });
    // Admin logout
    app.get('/admin-logout', (req, res) => {
    req.session.destroy(err => {    
    if (err) {
    console.error('Error destroying session:', err.message, err.stack);
    }
    res.redirect('/admin-login');
    });
    });

//admin-dashboard/media-uploads
    app.get('/admin-dashboard/media-uploads', requireAdmin, async (req, res) => {
    try {
        const allMedia = [];

        // --- Fetch Student uploads ---
        const studentsSnapshot = await db.collection('participants').get();
        for (const studentDoc of studentsSnapshot.docs) {
            const mediaSnapshot = await studentDoc.ref.collection('mediaUploads').get();
            mediaSnapshot.forEach(doc => {
                allMedia.push({
                    id: doc.id,
                    ...doc.data(),
                    uploadedBy: studentDoc.data().studentName || 'Unknown Student',
                    uploaderType: 'Student'
                });
            });
        }

        // --- Fetch School uploads ---
        const schoolsSnapshot = await db.collection('schools').get();
        for (const schoolDoc of schoolsSnapshot.docs) {
            const mediaSnapshot = await schoolDoc.ref.collection('mediaUploads').get();
            mediaSnapshot.forEach(doc => {
                allMedia.push({
                    id: doc.id,
                    ...doc.data(),
                    uploadedBy: schoolDoc.data().schoolName || 'Unknown School',
                    uploaderType: 'School'
                });
            });
        }

        // --- Fetch Trainer uploads ---
        const trainersSnapshot = await db.collection('trainers').get();
        for (const trainerDoc of trainersSnapshot.docs) {
            const mediaSnapshot = await trainerDoc.ref.collection('mediaUploads').get();
            mediaSnapshot.forEach(doc => {
                allMedia.push({
                    id: doc.id,
                    ...doc.data(),
                    uploadedBy: trainerDoc.data().trainerName || 'Unknown Trainer',
                    uploaderType: 'Trainer'
                });
            });
        }

        // Sort by uploadedAt (newest first)
        allMedia.sort((a, b) => {
            const dateA = a.uploadedAt?.toDate ? a.uploadedAt.toDate() : new Date(0);
            const dateB = b.uploadedAt?.toDate ? b.uploadedAt.toDate() : new Date(0);
            return dateB - dateA;
        });

        res.render('adminMediaDashboard', { mediaUploads: allMedia });

    } catch (error) {
        console.error('Error fetching all media uploads:', error);
        res.status(500).send('Error loading media uploads');
    }
});

    //traniner game zone
app.get('/trainer-gamezone', async (req, res) => {
    try {
        res.render('trainer-gamezone', { trainerName: 'Demo Trainer', trainerMobile: 'N/A' });
    } catch (error) {
        console.error('Error in trainer-gamezone route:', error.message);
        res.status(500).send('Error loading Trainer Game Zone.');
    }
});
// trainer add date
app.post('/trainer-dashboard/add-dates', async (req, res) => {
    try {
        const { trainerEmail, availableDates } = req.body;
        if (!trainerEmail || !availableDates) return res.status(400).send('Missing data');

        const trainerSnapshot = await db.collection('trainers').where('email', '==', trainerEmail).get();

        if (trainerSnapshot.empty) return res.status(404).send('Trainer not found');

        const trainerId = trainerSnapshot.docs[0].id;
        const trainerRef = db.collection('trainers').doc(trainerId);
        const trainerData = trainerSnapshot.docs[0].data();

        const existingDates = trainerData.availableDates || [];
        const newDates = Array.isArray(availableDates) ? availableDates : [availableDates];

        await trainerRef.update({
            availableDates: [...existingDates, ...newDates]
        });

        res.redirect(`/trainer-dashboard?username=${encodeURIComponent(trainerEmail)}`);
    } catch (err) {
        console.error('Error adding dates:', err);
        res.status(500).send('Internal Server Error');
    }
});

    //delete date for trainer 
  app.post('/trainer-dashboard/delete-date', async (req, res) => {
    try {
        const { trainerEmail, dateToDelete } = req.body;
        const username = req.query.username;

        if (!trainerEmail || trainerEmail !== username) {
            return res.redirect(`/trainer-dashboard?username=${encodeURIComponent(username)}&error=Invalid%20trainer%20email`);
        }

        if (!dateToDelete) {
            return res.redirect(`/trainer-dashboard?username=${encodeURIComponent(username)}&error=No%20date%20provided`);
        }

        const trainerSnapshot = await db.collection('trainers').where('email', '==', trainerEmail).get();

        if (trainerSnapshot.empty) {
            return res.redirect(`/trainer-dashboard?username=${encodeURIComponent(username)}&error=Trainer%20not%20found`);
        }

        const trainerId = trainerSnapshot.docs[0].id;
        const trainerRef = db.collection('trainers').doc(trainerId);
        const trainerData = trainerSnapshot.docs[0].data();

        const dateToDeleteStr = new Date(dateToDelete).toISOString().split('T')[0];
        if (!dateToDeleteStr) {
            return res.redirect(`/trainer-dashboard?username=${encodeURIComponent(username)}&error=Invalid%20date%20format`);
        }

        const updatedDates = (trainerData.availableDates || []).filter(date => {
            if (!date) return true;

            // If stored as Firestore Timestamp
            if (date.toDate) {
                return date.toDate().toISOString().split('T')[0] !== dateToDeleteStr;
            }

            // If stored as string
            const dateStr = new Date(date).toISOString().split('T')[0];
            return dateStr !== dateToDeleteStr;
        });

        await trainerRef.update({
            availableDates: updatedDates
        });

        return res.redirect(`/trainer-dashboard?username=${encodeURIComponent(username)}&success=Date%20deleted%20successfully`);
    } catch (error) {
        console.error('Error deleting date:', error);
        return res.redirect(`/trainer-dashboard?username=${encodeURIComponent(req.query.username)}&error=Error%20deleting%20date`);
    }
});

    //uploaded media for trainer
   app.post('/trainer-dashboard/upload-media', upload, async (req, res) => {
    try {
        const { trainerEmail, mediaDescription, mediaLink } = req.body;
        const photos = req.files['photos'] || [];
        const videos = req.files['videos'] || [];
        // const zipFiles = req.files['zipFiles'] || []; // Add ZIP file handling

        if (!trainerEmail) {
            return res.status(400).render('trainerDashboard', {
                trainerName: 'Unknown',
                email: '',
                city: '',
                district: '',
                profession: '',
                assignedSchools: [],
                availableDates: [],
                schoolsInDistrict: [],
                mediaUploads: [],
                approvedZips: [],
                error: 'Trainer email is required',
                success: null,
                trainerData: {}
            });
        }

        // Fetch trainer data
        const trainerSnapshot = await db.collection('trainers')
            .where('email', '==', trainerEmail)
            .get();
        if (trainerSnapshot.empty) {
            return res.status(404).render('trainerDashboard', {
                trainerName: 'Unknown',
                email: trainerEmail,
                city: '',
                district: '',
                profession: '',
                assignedSchools: [],
                availableDates: [],
                schoolsInDistrict: [],
                mediaUploads: [],
                approvedZips: [],
                error: 'Trainer not found',
                success: null,
                trainerData: {}
            });
        }

        const trainerId = trainerSnapshot.docs[0].id;
        const trainerData = trainerSnapshot.docs[0].data();

        if (!trainerData.isApproved) {
            return res.status(403).render('trainerDashboard', {
                trainerName: trainerData.trainerName || 'Unknown',
                email: trainerEmail,
                city: trainerData.city || '',
                district: trainerData.district || trainerData.city || '',
                profession: trainerData.profession || '',
                assignedSchools: [],
                availableDates: trainerData.availableDates || [],
                schoolsInDistrict: [],
                mediaUploads: [],
                approvedZips: [],
                error: 'You must complete the Train the Trainer program to upload media.',
                success: null,
                trainerData
            });
        }

        if (mediaDescription && mediaDescription.length > 1000) {
            return res.status(400).render('trainerDashboard', {
                trainerName: trainerData.trainerName || 'Unknown',
                email: trainerEmail,
                city: trainerData.city || '',
                district: trainerData.district || trainerData.city || '',
                profession: trainerData.profession || '',
                assignedSchools: [],
                availableDates: trainerData.availableDates || [],
                schoolsInDistrict: [],
                mediaUploads: [],
                approvedZips: [],
                error: 'Media description cannot exceed 1000 characters',
                success: null,
                trainerData
            });
        }

        // Explicitly set the bucket name
        const bucket = admin.storage().bucket('beinglawful-ee5a4.firebasestorage.app');
        const [exists] = await bucket.exists();
        if (!exists) {
            throw new Error('The specified bucket does not exist. Please check your Firebase configuration.');
        }

        const mediaUploads = [];
        const zipUploads = [];

        // Handle photos
        for (const file of photos) {
            const fileName = `media/trainers/${trainerId}/photos/${Date.now()}_${file.originalname}`;
            const fileUpload = bucket.file(fileName);
            await fileUpload.save(file.buffer, {
                metadata: { contentType: file.mimetype }
            });
            const [url] = await fileUpload.getSignedUrl({
                action: 'read',
                expires: '03-09-2491'
            });
            mediaUploads.push({
                url,
                type: 'image',
                path: fileName,
                description: mediaDescription || '',
                link: mediaLink || '',
                uploadedAt: admin.firestore.FieldValue.serverTimestamp()
            });
        }

        // Handle videos
        for (const file of videos) {
            const fileName = `media/trainers/${trainerId}/videos/${Date.now()}_${file.originalname}`;
            const fileUpload = bucket.file(fileName);
            await fileUpload.save(file.buffer, {
                metadata: { contentType: file.mimetype }
            });
            const [url] = await fileUpload.getSignedUrl({
                action: 'read',
                expires: '03-09-2491'
            });
            mediaUploads.push({
                url,
                type: 'video',
                path: fileName,
                description: mediaDescription || '',
                link: mediaLink || '',
                uploadedAt: admin.firestore.FieldValue.serverTimestamp()
            });
        }

        // Handle ZIP files
        for (const file of zipFiles) {
            const fileName = `zipfiles/trainers/${trainerId}/${Date.now()}_${file.originalname}`;
            const fileUpload = bucket.file(fileName);
            await fileUpload.save(file.buffer, {
                metadata: { contentType: file.mimetype }
            });
            const [url] = await fileUpload.getSignedUrl({
                action: 'read',
                expires: '03-09-2491'
            });
            zipUploads.push({
                trainerId,
                fileName: file.originalname,
                url,
                path: fileName,
                approved: false, // Initially not approved
                uploadedAt: admin.firestore.FieldValue.serverTimestamp(),
                description: mediaDescription || '',
                link: mediaLink || ''
            });
        }

        // Save media uploads to Firestore
        if (mediaUploads.length > 0) {
            const batch = db.batch();
            mediaUploads.forEach(upload => {
                const mediaRef = db.collection('trainers').doc(trainerId).collection('mediaUploads').doc();
                batch.set(mediaRef, upload);
            });
            await batch.commit();
        }

        // Save ZIP uploads to Firestore
        if (zipUploads.length > 0) {
            const batch = db.batch();
            zipUploads.forEach(upload => {
                const zipRef = db.collection('trainerZips').doc();
                batch.set(zipRef, upload);
            });
            await batch.commit();
        }

        // Fetch updated trainer data
        const updatedTrainerSnapshot = await db.collection('trainers').doc(trainerId).get();
        const updatedTrainerData = updatedTrainerSnapshot.data();

        // Fetch assigned schools
        const schoolsSnapshot1 = await db.collection('schools')
            .where('isApproved', '==', true)
            .where('trainerId1', '==', trainerId)
            .get();
        const schoolsSnapshot2 = await db.collection('schools')
            .where('isApproved', '==', true)
            .where('trainerId2', '==', trainerId)
            .get();
        const assignedSchools = [...schoolsSnapshot1.docs, ...schoolsSnapshot2.docs].map(doc => {
            const data = doc.data();
            return {
                schoolName: data.schoolName || 'N/A',
                city: data.city || 'N/A',
                district: data.district || data.city || 'N/A',
                eventDate: data.eventDate ? data.eventDate.toDate().toLocaleDateString('en-US', {
                    year: 'numeric',
                    month: 'long',
                    day: 'numeric'
                }) : 'Not assigned',
                resourcesConfirmed: data.resourcesConfirmed || false,
                selectedResources: data.selectedResources || [],
                trainerRole: data.trainerId1 === trainerId ? 'Trainer 1' : 'Trainer 2'
            };
        });

        // Fetch schools in the trainer's district
        const schoolsInDistrictSnapshot = await db.collection('schools')
            .where('district', '==', updatedTrainerData.district || updatedTrainerData.city)
            .where('isApproved', '==', true)
            .get();
        const schoolsInDistrict = schoolsInDistrictSnapshot.docs.map(doc => {
            const data = doc.data();
            return {
                id: doc.id,
                schoolName: data.schoolName || 'N/A',
                schoolEmail: data.schoolEmail || 'N/A',
                city: data.city || 'N/A',
                district: data.district || 'N/A',
                eventDate: data.eventDate ? data.eventDate.toDate().toLocaleDateString('en-US', {
                    year: 'numeric',
                    month: 'long',
                    day: 'numeric'
                }) : 'Not assigned',
                assignedTrainerId: data.trainerId1 || data.trainerId2 || null,
                trainerRole: data.trainerId1 === trainerId ? 'Trainer 1' : data.trainerId2 === trainerId ? 'Trainer 2' : null
            };
        });

        // Fetch media uploads
        const mediaSnapshot = await db.collection('trainers').doc(trainerId).collection('mediaUploads').get();
        const mediaUploadsData = mediaSnapshot.docs.map(doc => ({
            ...doc.data(),
            uploadedAt: doc.data().uploadedAt ? doc.data().uploadedAt.toDate() : null
        }));

        // Fetch approved ZIP files
        const zipSnapshot = await db.collection('trainerZips')
            .where('trainerId', '==', trainerId)
            .where('approved', '==', true)
            .get();
        const approvedZips = zipSnapshot.docs.map(doc => ({
            ...doc.data(),
            downloadUrl: `/zipfiles/${doc.data().fileName}`,
            uploadedAt: doc.data().uploadedAt ? doc.data().uploadedAt.toDate() : null
        }));

        res.render('trainerDashboard', {
            trainerName: updatedTrainerData.trainerName || 'Unknown',
            email: trainerEmail,
            city: updatedTrainerData.city || '',
            district: updatedTrainerData.district || updatedTrainerData.city || '',
            profession: updatedTrainerData.profession || '',
            assignedSchools,
            availableDates: updatedTrainerData.availableDates || [],
            schoolsInDistrict,
            mediaUploads: mediaUploadsData,
            approvedZips,
            error: null,
            success: 'Media and/or ZIP files uploaded successfully',
            trainerData: updatedTrainerData
        });
    } catch (error) {
        console.error('Error in trainer-dashboard/upload-media route:', error.message, error.stack);
        res.status(500).render('trainerDashboard', {
            trainerName: 'Unknown',
            email: trainerEmail || '',
            city: '',
            district: '',
            profession: '',
            assignedSchools: [],
            availableDates: [],
            schoolsInDistrict: [],
            mediaUploads: [],
            approvedZips: [],
            error: `Error uploading media: ${error.message}`,
            success: null,
            trainerData: {}
        });
    }
});

app.get('/api/trainers', async (req, res) => {
  try {
    const trainersSnapshot = await db.collection('trainers').get();

    const trainers = await Promise.all(trainersSnapshot.docs.map(async (doc) => {
      const data = doc.data();
      const trainerId = doc.id;

      // Fetch completed schools for this trainer
      const completedSnapshot1 = await db.collection('schools')
        .where('trainerId1', '==', trainerId)
        .where('isCompleted', '==', true)
        .get();

      const completedSnapshot2 = await db.collection('schools')
        .where('trainerId2', '==', trainerId)
        .where('isCompleted', '==', true)
        .get();

      const completedSchoolsCount = completedSnapshot1.docs.length + completedSnapshot2.docs.length;

      return {
        trainerName: data.trainerName || 'N/A',
        city: data.city || 'N/A',
        district: data.district || 'N/A',
        profession: data.profession || 'N/A',
        profilePhoto: data.profilePhoto || null, // Include profile photo URL
        completedSchoolsCount // Count of completed schools
      };
    }));

    // Sort trainers by completedSchoolsCount (descending) and then by trainerName (alphabetically)
    trainers.sort((a, b) => {
      if (b.completedSchoolsCount !== a.completedSchoolsCount) {
        return b.completedSchoolsCount - a.completedSchoolsCount;
      }
      return a.trainerName.localeCompare(b.trainerName);
    });

    // Render the ourtrainers.ejs template with trainers data
    res.render('ourtrainers', { trainers });
  } catch (error) {
    console.error('Error fetching trainers:', error);
    res.status(500).send('Error fetching trainers: ' + error.message);
  }
});

app.get('/ourtrainers', async (req, res) => {
  try {
    const response = await axios.get('http://localhost:3000/api/trainers');
    res.render('ourtrainers', { trainers: response.data });
  } catch (error) {
    console.error('Error fetching trainers for /ourtrainers:', error);
    res.status(500).send('Error rendering trainers page');
  }
});
// Trainer Dashboard Route
app.get('/trainer-dashboard', async (req, res) => {
    try {
        const trainerEmail = req.query.username;
        if (!trainerEmail) {
            return res.status(400).render('trainerDashboard', {
                trainerName: 'Unknown',
                email: '',
                city: '',
                district: '',
                profession: '',
                assignedSchools: [],
                completedSchools: [],
                availableDates: [],
                schoolsInDistrict: [],
                mediaUploads: [],
                approvedZips: [],
                error: 'Trainer email is required',
                success: null,
                trainerData: {}
            });
        }

        // Fetch trainer data
        const trainerSnapshot = await db.collection('trainers')
            .where('email', '==', trainerEmail)
            .get();

        if (trainerSnapshot.empty) {
            return res.status(404).render('trainerDashboard', {
                trainerName: 'Unknown',
                email: trainerEmail,
                city: '',
                district: '',
                profession: '',
                assignedSchools: [],
                completedSchools: [],
                availableDates: [],
                schoolsInDistrict: [],
                mediaUploads: [],
                approvedZips: [],
                error: 'Trainer not found',
                success: null,
                trainerData: {}
            });
        }

        const trainerId = trainerSnapshot.docs[0].id;
        const trainerData = trainerSnapshot.docs[0].data();

        // Fetch assigned schools
        const schoolsSnapshot1 = await db.collection('schools')
            .where('isApproved', '==', true)
            .where('trainerId1', '==', trainerId)
            .get();

        const schoolsSnapshot2 = await db.collection('schools')
            .where('isApproved', '==', true)
            .where('trainerId2', '==', trainerId)
            .get();

        const assignedSchools = await Promise.all([...schoolsSnapshot1.docs, ...schoolsSnapshot2.docs].map(async doc => {
            const data = doc.data();
            let coordinatorName = 'N/A';
            let coordinatorNumber = 'N/A';

            // Fetch coordinator data if coordinatorId exists
            if (data.coordinatorId && data.coordinatorId !== 'N/A') {
                const coordinatorSnapshot = await db.collection('coordinators')
                    .doc(data.coordinatorId)
                    .get();
                
                if (coordinatorSnapshot.exists) {
                    const coordinatorData = coordinatorSnapshot.data();
                    coordinatorName = coordinatorData.name || 'N/A';
                    coordinatorNumber = coordinatorData.number || 'N/A';
                }
            }

            return {
                schoolName: data.schoolName || 'N/A',
                city: data.city || 'N/A',
                district: data.district || data.city || 'N/A',
                eventDate: data.eventDate ? data.eventDate.toDate().toLocaleDateString('en-US', {
                    year: 'numeric',
                    month: 'long',
                    day: 'numeric'
                }) : 'Not assigned',
                resourcesConfirmed: data.resourcesConfirmed || false,
                trainerRole: data.trainerId1 === trainerId ? 'Trainer 1' : 'Trainer 2',
                schoolPhoneNumber: data.schoolPhoneNumber || 'N/A',
                civicsTeacherNumber: data.civicsTeacherNumber || 'N/A',
                principalNumber: data.principalNumber || 'N/A',
                principalEmail: data.principalEmail || 'N/A',
                schoolCode: data.schoolCode || data.schoolNumber || 'N/A',
                coordinatorId: data.coordinatorId || 'N/A',
                coordinatorName, 
                coordinatorNumber 
            };
        }));
console.log('Assigned Schools:', assignedSchools);
        // Fetch completed schools
        const completedSnapshot1 = await db.collection('schools')
            .where('trainerId1', '==', trainerId)
            .where('isCompleted', '==', true)
            .get();

        const completedSnapshot2 = await db.collection('schools')
            .where('trainerId2', '==', trainerId)
            .where('isCompleted', '==', true)
            .get();

        const completedSchools = await Promise.all([...completedSnapshot1.docs, ...completedSnapshot2.docs].map(async doc => {
            const data = doc.data();
            let coordinatorName = 'N/A';
            let coordinatorNumber = 'N/A';

            // Fetch coordinator data if coordinatorId exists
            if (data.coordinatorId && data.coordinatorId !== 'N/A') {
                const coordinatorSnapshot = await db.collection('coordinators')
                    .doc(data.coordinatorId)
                    .get();
                
                if (coordinatorSnapshot.exists) {
                    const coordinatorData = coordinatorSnapshot.data();
                    coordinatorName = coordinatorData.name || 'N/A';
                    coordinatorNumber = coordinatorData.number || 'N/A';
                }
            }

            return {
                schoolName: data.schoolName || 'N/A',
                city: data.city || 'N/A',
                district: data.district || data.city || 'N/A',
                eventDate: data.eventDate ? data.eventDate.toDate().toLocaleDateString('en-IN') : 'Not assigned',
                completionDate: data.completionDate ? data.completionDate.toDate().toLocaleDateString('en-IN') : 'N/A',
                remarks: data.remarks || 'N/A',
                coordinatorId: data.coordinatorId || 'N/A',
                coordinatorName, // Add coordinator name
                coordinatorNumber // Add coordinator number
            };
        }));

        // Fetch schools in trainer's district
        const schoolsInDistrictSnapshot = await db.collection('schools')
            .where('district', '==', trainerData.district || trainerData.city)
            .where('isApproved', '==', true)
            .get();

        const schoolsInDistrict = schoolsInDistrictSnapshot.docs.map(doc => {
            const data = doc.data();
            return {
                id: doc.id,
                schoolName: data.schoolName || 'N/A',
                schoolEmail: data.schoolEmail || 'N/A',
                city: data.city || 'N/A',
                district: data.district || 'N/A',
                eventDate: data.eventDate ? data.eventDate.toDate().toLocaleDateString('en-US', {
                    year: 'numeric',
                    month: 'long',
                    day: 'numeric'
                }) : 'Not assigned',
                assignedTrainerId: data.trainerId1 || data.trainerId2 || null,
                trainerRole: data.trainerId1 === trainerId ? 'Trainer 1' : data.trainerId2 === trainerId ? 'Trainer 2' : null
            };
        });

        // Fetch media uploads
        const mediaSnapshot = await db.collection('trainers').doc(trainerId).collection('mediaUploads').get();
        const mediaUploads = mediaSnapshot.docs.map(doc => ({
            ...doc.data(),
            uploadedAt: doc.data().uploadedAt ? doc.data().uploadedAt.toDate() : null
        }));

        // Fetch approved ZIP files for trainers
        const zipSnapshot = await db.collection('trainerZips')
            .where('trainerId', '==', trainerId)
            .where('approved', '==', true)
            .get();

        const approvedZips = zipSnapshot.docs.map(doc => {
            const data = doc.data();
            return {
                ...data,
                downloadUrl: `/zipfiles/${data.fileName}`,
                uploadedAt: data.uploadedAt ? data.uploadedAt.toDate() : null
            };
        });
         
        // Render Trainer Dashboard
        res.render('trainerDashboard', {
            trainerName: trainerData.trainerName || 'Unknown',
            email: trainerEmail,
            city: trainerData.city || '',
            district: trainerData.district || trainerData.city || '',
            profession: trainerData.profession || '',
            assignedSchools,
            completedSchools,
            availableDates: trainerData.availableDates || [],
            schoolsInDistrict,
            mediaUploads,
            approvedZips,
            error: req.query.error || null,
            success: req.query.success || null,
            trainerData,
        });

    } catch (error) {
        console.error('Error in trainer-dashboard route:', error.message, error.stack);
        res.status(500).render('trainerDashboard', {
            trainerName: 'Unknown',
            email: req.query.username || '',
            city: '',
            district: '',
            profession: '',
            assignedSchools: [],
            completedSchools: [],
            availableDates: [],
            schoolsInDistrict: [],
            mediaUploads: [],
            approvedZips: [],
            error: `Error loading dashboard: ${error.message}`,
            success: null,
            trainerData: {}
        });
    }
    console.log('assigned schools',assignedSchools)
});


// Updated profile route
app.get('/trainer-dashboard/profile/:username', async (req, res) => {
  try {
    const email = decodeURIComponent(req.params.username);
    if (!email) {
      return res.status(400).render('profile', {
        trainerName: 'Unknown',
        email: '',
        city: '',
        district: '',
        profession: '',
        trainerData: {},
        error: 'Invalid or missing username',
        success: null,
      });
    }

    const trainerSnapshot = await db.collection('trainers').where('email', '==', email).get();
    if (trainerSnapshot.empty) {
      return res.status(404).render('profile', {
        trainerName: 'Unknown',
        email: email,
        city: '',
        district: '',
        profession: '',
        trainerData: {},
        error: 'Trainer not found',
        success: null,
      });
    }

    const trainerData = trainerSnapshot.docs[0].data();
    res.render('profile', {
      trainerName: trainerData.trainerName || 'Unknown',
      email: trainerData.email || email,
      city: trainerData.city || '',
      district: trainerData.district || trainerData.city || '',
      profession: trainerData.profession || '',
      trainerData: trainerData,
      error: null,
      success: null,
    });
  } catch (error) {
    console.error('Error in trainer-dashboard/profile route:', error.message, error.stack);
    res.status(500).render('profile', {
      trainerName: 'Unknown',
      email: req.params.username || '',
      city: '',
      district: '',
      profession: '',
      trainerData: {},
      error: `Error loading profile: ${error.message}`,
      success: null,
    });
  }
});
// Configure Multer for file uploads
const storages = multer.memoryStorage();
const trainerUploads = multer({ storages }).single('profilePhoto');

app.post('/trainer-dashboard/upload-profile-photo', trainerUploads, async (req, res) => {
    try {
        const trainerEmail = req.body.trainerEmail;
        if (!trainerEmail) {
            throw new Error('Trainer email is required');
        }

        if (!req.file) {
            throw new Error('No file uploaded');
        }

        // Fetch trainer from Firestore
        const trainerSnapshot = await db.collection('trainers')
            .where('email', '==', trainerEmail)
            .get();

        if (trainerSnapshot.empty) {
            throw new Error('Trainer not found');
        }

        const trainerId = trainerSnapshot.docs[0].id;
        const trainerDoc = trainerSnapshot.docs[0];

        // Generate a unique file name
        const fileName = `profile-photos/${trainerId}-${Date.now()}-${req.file.originalname}`;
        const file = bucket.file(fileName);

        // Validate file type (optional, for security)
        const allowedTypes = ['image/jpeg', 'image/png', 'image/gif'];
        if (!allowedTypes.includes(req.file.mimetype)) {
            throw new Error('Invalid file type. Only JPEG, PNG, and GIF are allowed.');
        }

        // Create a stream to upload the file to Firebase Storage
        const stream = file.createWriteStream({
            metadata: {
                contentType: req.file.mimetype
            }
        });

        stream.on('error', (error) => {
            console.error('Error uploading to Firebase Storage:', error);
            throw new Error('Error uploading file to storage');
        });

        stream.on('finish', async () => {
            try {
                // Get a signed URL instead of making the file public
                const [photoUrl] = await file.getSignedUrl({
                    action: 'read',
                    expires: '03-09-2491' // Far-future expiration
                });
                console.log('Generated Photo URL:', photoUrl);

                // Delete old profile photo if it exists
                const oldPhotoUrl = trainerDoc.data().profilePhoto;
                if (oldPhotoUrl) {
                    try {
                        // Extract file name from URL
                        const oldFileName = oldPhotoUrl.includes(bucket.name)
                            ? oldPhotoUrl.split(`${bucket.name}/`)[1]
                            : null;
                        if (oldFileName) {
                            await bucket.file(oldFileName).delete();
                            console.log(`Deleted old profile photo: ${oldFileName}`);
                        } else {
                            console.warn('Could not extract old file name from URL:', oldPhotoUrl);
                        }
                    } catch (error) {
                        console.error(`Error deleting old profile photo: ${error.message}`);
                    }
                }

                // Update trainer's Firestore document with the photo URL
                await db.collection('trainers').doc(trainerId).update({
                    profilePhoto: photoUrl,
                    updatedAt: admin.firestore.FieldValue.serverTimestamp()
                });

                // Redirect with success message
                res.redirect(`/trainer-dashboard?username=${trainerEmail}&success=Profile photo uploaded successfully`);
            } catch (error) {
                console.error('Error in post-upload processing:', error);
                throw new Error('Error processing file upload');
            }
        });

        // Write the file buffer to the stream
        stream.end(req.file.buffer);
    } catch (error) {
        console.error('Error in profile photo upload:', error.message, error.stack);
        res.redirect(`/trainer-dashboard?username=${req.body.trainerEmail || req.query.username || ''}&error=${encodeURIComponent(error.message)}`);
    }
});
 
app.post('/trainer-login', async (req, res) => {
            const { username, password } = req.body;
            try {
                if (!username || !password) {
                    return res.render('login', { error: 'Username and password are required.' });
                }

                const snapshot = await db.collection('trainers')
                    .where('email', '==', username)
                    .get();
                if (snapshot.empty) {
                    return res.render('login', { error: 'Invalid username or password.' });
                }

                const trainer = snapshot.docs[0].data();
                if (password !== trainer.mobileNumber) {
                    return res.render('login', { error: 'Invalid username or password.' });
                }

                res.redirect(`/trainer-dashboard?username=${encodeURIComponent(username)}`);
            } catch (error) {
                console.error('Error during trainer login:', error.message, error.stack);
                res.render('login', { error: 'Login failed. Try again later.' });
            }
        });

  app.get('/trainer-participation', async (req, res) => {
    try {
        const schoolsSnapshot = await db.collection('schools')
            .where('isApproved', '==', true)
            .get();
        const schools = schoolsSnapshot.docs.map(doc => doc.data().schoolName);

        const trainersSnapshot = await db.collection('trainers')
            .where('isApproved', '==', true) // Filter for approved trainers
            .get();
        const trainers = trainersSnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));

        res.render('trainerParticipation', { schools, trainers, errors: null });
    } catch (error) {
        console.error('Error in trainer-participation route:', error.message, error.stack);
        res.status(500).send('Error loading trainer participation form.');
    }
});
   
app.post('/trainer-participate', [
    body('trainerName').trim().notEmpty().withMessage('Trainer name is required'),
    body('email').trim().isEmail().withMessage('Invalid email address'),
    body('city').trim().notEmpty().withMessage('City is required'),
    body('district').trim().notEmpty().withMessage('District is required'),
    body('profession').trim().notEmpty().withMessage('Profession is required'),
    body('mobileNumber').trim().matches(/^\d{10}$/).withMessage('Mobile number must be 10 digits'),
    body('whatsappNumber').trim().matches(/^\d{10}$/).withMessage('WhatsApp number must be 10 digits'),
    body('referenceName').trim().notEmpty().withMessage('Reference name is required'),
    body('otherProfession').trim().custom((value, { req }) => {
        if (req.body.profession === 'Other' && !value) {
            throw new Error('Other profession is required when profession is "Other"');
        }
        return true;
    }),
], async (req, res) => {
    try {
        const validationErrors = validationResult(req);
        const errors = validationErrors.isEmpty() ? [] : validationErrors.array();

        const {
            trainerName,
            email,
            city,
            district,
            profession,
            mobileNumber,
            whatsappNumber,
            referenceName,
            otherProfession,
        } = req.body;

        // Determine the profession to store
        const finalProfession = profession === 'Other' ? otherProfession : profession;

        // Validate that finalProfession is not empty
        if (!finalProfession) {
            errors.push({ param: 'profession', msg: 'Profession cannot be empty' });
        }

        // Check for duplicates
        const duplicateChecks = [
            { field: 'email', value: email, message: 'Email already exists' },
            { field: 'mobileNumber', value: mobileNumber, message: 'Mobile number already exists' },
            { field: 'whatsappNumber', value: whatsappNumber, message: 'WhatsApp number already exists' },
        ];

        for (const check of duplicateChecks) {
            const snapshot = await db.collection('trainers').where(check.field, '==', check.value).get();
            if (!snapshot.empty) {
                errors.push({ param: check.field, msg: check.message });
            }
        }

        if (errors.length > 0) {
            return await renderFormWithErrors(res, errors);
        }

        // Store data in Firebase
        await db.collection('trainers').add({
            trainerName,
            email,
            city,
            district,
            profession: finalProfession, // Store finalProfession
            mobileNumber,
            whatsappNumber,
            referenceName,
            registeredAt: admin.firestore.FieldValue.serverTimestamp(),
            isApproved: false,
        });

        // Uncomment and implement email sending if needed
        /*
        await sendEmail(
            email,
            emailTemplates.trainerRegistration.subject,
            emailTemplates.trainerRegistration.text(trainerName, email, mobileNumber),
            emailTemplates.trainerRegistration.html(trainerName, email, mobileNumber)
        );
        */

        // Render confirmation page
        res.render('confirmation', {
            schoolEmail: email,
            principalNumber: mobileNumber,
        });
    } catch (error) {
        console.error('Error in trainer-participate route:', error.message, error.stack);
        await renderFormWithErrors(res, [{ msg: 'Internal server error. Please try again later.' }]);
    }
});

// Helper function to render form with data and errors
async function renderFormWithErrors(res, errors) {
    const schoolsSnapshot = await db.collection('schools').where('isApproved', '==', true).get();
    const schools = schoolsSnapshot.docs.map(doc => doc.data().schoolName);

    const trainersSnapshot = await db.collection('trainers').get();
    const trainers = trainersSnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));

    res.status(400).render('trainerParticipation', { schools, trainers, errors });
}

    // Assign trainer to school
    app.post('/admin/assign-trainer', requireAdmin, async (req, res) => {
    try {
    const { schoolId, trainerId } = req.body;
    if (!schoolId || !trainerId) {
    return res.redirect('/admin-dashboard?error=School ID and Trainer ID are required');
    }

    const schoolRef = db.collection('schools').doc(schoolId);
    const schoolDoc = await schoolRef.get();
    if (!schoolDoc.exists) {
    return res.redirect('/admin-dashboard?error=School not found');
    }

    const trainerRef = db.collection('trainers').doc(trainerId);
    const trainerDoc = await trainerRef.get();
    if (!trainerDoc.exists) {
    return res.redirect('/admin-dashboard?error=Trainer not found');
    }

    await schoolRef.update({ assignedTrainerId: trainerId });
    res.redirect('/admin-dashboard?success=Trainer assigned successfully');
    } catch (error) {
    console.error('Error in admin/assign-trainer route:', error.message, error.stack);
    res.redirect('/admin-dashboard?error=Error assigning trainer');
    }
    });

    // Logout
    app.get('/logout', (req, res) => {
    req.session.destroy(err => {
    if (err) {
    console.error('Error destroying session:', err.message, err.stack);
    }
    res.redirect('/login');
    });
    });


    // logistic Dashboard route
// Logistic Dashboard Route
app.get('/logistic-dashboard', async (req, res) => {
  try {
    const snapshot = await db
      .collection('schools')
      .where('isApproved', '==', true)
      .get();

    const schools = [];

    snapshot.forEach(doc => {
      const data = doc.data();
      const rawDate = data.eventDate?.toDate();

      schools.push({
        id: doc.id,
        name: data.schoolName || 'N/A',
        civicsSirNumber: data.civicsTeacherNumber || 'N/A',
        schoolPhoneNumber: data.schoolPhoneNumber || 'N/A',
        principalNumber: data.principalNumber || 'N/A',
        city: data.city || 'N/A',
        district: data.district || 'N/A',
        agency: data.agency || '',
        deliveryMode: data.deliveryMode || 'pending',
        deliveryNumber: data.deliveryNumber || '',
        rawEventDate: rawDate,
        eventDate: rawDate ? rawDate.toLocaleDateString('en-IN') : 'N/A',
        status: data.status || 'pending' // Use status field instead of isCompleted
      });
    });

    // Sort by rawEventDate (null-safe)
    schools.sort((a, b) => {
      if (!a.rawEventDate) return 1;
      if (!b.rawEventDate) return -1;
      return a.rawEventDate - b.rawEventDate;
    });

    res.render('logistic-dashboard', { schools });
  } catch (error) {
    console.error('Error fetching school data:', error);
    res.status(500).send('Internal Server Error');
  }
});

// Route to fetch financial overview
app.get('/financial', async (req, res) => {
  try {
    const schoolSnapshot = await db.collection('schools').where('isApproved', '==', true).get();
    const schools = [];
    const paymentSnapshot = await db.collection('payments').get();
    const payments = {};
    const trainerSnapshot = await db.collection('trainers').get();
    const trainers = [];
    const workshopSnapshot = await db.collection('workshopSummaries').get();
    const workshops = {};

    // Process trainers data
    trainerSnapshot.forEach(doc => {
      const data = doc.data();
      trainers.push({
        id: doc.id,
        name: data.trainerName || 'N/A'
      });
    });

    // Process payments data
    paymentSnapshot.forEach(doc => {
      const data = doc.data();
      const totalExpenses = (Number(data.kitCharges) || 0) +
                           (Number(data.contentRoyalty) || 0) +
                           (Number(data.bookCharges) || 0) +
                           (Number(data.trainerRemunerationAmount) || 0) +
                           (Number(data.travellingAllowance) || 0) +
                           (Number(data.petrolDiesel) || 0);
      payments[doc.id] = {
        workshopFees: Number(data.workshopFees) || 11000,
        invoiceStatus: data.invoiceStatus || 'N/A',
        kitCharges: Number(data.kitCharges) || 0,
        kitOutDate: data.kitOutDate?.toDate() ? data.kitOutDate.toDate().toLocaleDateString('en-IN') : 'N/A',
        kitReceivedBy: data.kitReceivedBy || 'N/A',
        status: data.status || 'N/A',
        contentRoyalty: Number(data.contentRoyalty) || 0,
        bookCharges: Number(data.bookCharges) || 800,
        trainerName1: data.trainerName1 || 'N/A',
        trainerName2: data.trainerName2 || 'N/A',
        trainerRemunerationAmount: Number(data.trainerRemunerationAmount) || 0,
        trainerTotalCost: Number(data.trainerTotalCost) || 0,
        travellingAllowance: Number(data.travellingAllowance) || 0,
        petrolDiesel: Number(data.petrolDiesel) || 0,
        paymentDate: data.paymentDate?.toDate() ? data.paymentDate.toDate().toLocaleDateString('en-IN') : 'N/A',
        bltAccount: data.bltAccount || 'N/A',
        paymentMethod: data.paymentMethod || 'CASH',
        totalExpenses,
        total: (Number(data.workshopFees) || 11000) - totalExpenses,
        financialStatus: data.financialStatus || 'Pending',
        kitPaymentStatus: data.kitPaymentStatus || 'Unpaid',
        trainerRemunerationStatus: data.trainerRemunerationStatus || 'Pending',
        upiId: data.upiId || 'N/A',
        transactionId: data.transactionId || 'N/A'
      };
    });

    // Process workshop summaries data
    workshopSnapshot.forEach(doc => {
      const data = doc.data();
      workshops[data.schoolName] = {
        financialStatus: data.financialStatus && data.financialStatus !== 'N/A' ? data.financialStatus : 'Pending',
        kitPaymentStatus: data.kitPaymentStatus && data.kitPaymentStatus !== 'N/A' ? data.kitPaymentStatus : 'Unpaid',
        trainerRemunerationStatus: data.trainerRemunerationStatus && data.trainerRemunerationStatus !== 'N/A' ? data.trainerRemunerationStatus : 'Pending',
        paymentMode: data.paymentMode || 'CASH',
        upiId: data.upiId || 'N/A',
        transactionId: data.transactionId || 'N/A',
        trainer1: data.trainer1 || 'N/A',
        trainer2: data.trainer2 || 'N/A'
      };
    });

    // Process schools data
    schoolSnapshot.forEach(doc => {
      const data = doc.data();
      const schoolId = doc.id;
      const paymentData = payments[schoolId] || {};
      const workshopData = workshops[data.schoolName] || {};
      const rawEventDate = data.eventDate?.toDate();

      schools.push({
        schoolId,
        schoolName: data.schoolName || 'N/A',
        eventDate: rawEventDate ? rawEventDate.toLocaleDateString('en-IN') : 'N/A',
        city: data.city || 'N/A',
        district: data.district || 'N/A',
        workshopFees: paymentData.workshopFees || 11000,
        invoiceStatus: paymentData.invoiceStatus || 'N/A',
        kitCharges: paymentData.kitCharges || 0,
        kitOutDate: paymentData.kitOutDate || 'N/A',
        kitReceivedBy: paymentData.kitReceivedBy || 'N/A',
        status: paymentData.status || 'N/A',
        contentRoyalty: paymentData.contentRoyalty || 0,
        bookCharges: paymentData.bookCharges || 800,
        trainerName1: paymentData.trainerName1 || workshopData.trainer1 || 'N/A',
        trainerName2: paymentData.trainerName2 || workshopData.trainer2 || 'N/A',
        trainerRemunerationAmount: paymentData.trainerRemunerationAmount || 0,
        trainerTotalCost: paymentData.trainerTotalCost || 0,
        travellingAllowance: paymentData.travellingAllowance || 0,
        petrolDiesel: paymentData.petrolDiesel || 0,
        paymentDate: paymentData.paymentDate || 'N/A',
        bltAccount: paymentData.bltAccount || 'N/A',
        paymentMethod: workshopData.paymentMode || paymentData.paymentMethod || 'CASH',
        totalExpenses: paymentData.totalExpenses || 0,
        total: paymentData.total || 0,
        financialStatus: workshopData.financialStatus || paymentData.financialStatus || 'Pending',
        kitPaymentStatus: workshopData.kitPaymentStatus || paymentData.kitPaymentStatus || 'Unpaid',
        trainerRemunerationStatus: workshopData.trainerRemunerationStatus || paymentData.trainerRemunerationStatus || 'Pending',
        upiId: workshopData.upiId || paymentData.upiId || 'N/A',
        transactionId: workshopData.transactionId || paymentData.transactionId || 'N/A'
      });
    });

    // Log schools data for debugging
    console.log('Schools data:', schools.map(s => ({
      schoolId: s.schoolId,
      trainerRemunerationAmount: s.trainerRemunerationAmount,
      trainerTotalCost: s.trainerTotalCost,
      totalExpenses: s.totalExpenses,
      total: s.total
    })));

    res.render('financial', { schools, trainers });
  } catch (error) {
    console.error('Error fetching school, payment, trainer, or workshop data:', error);
    res.status(500).send('Internal Server Error');
  }
});
app.get('/other', (req, res) => {
  res.render('other', {});
});

// Route to update financial details
app.post('/financial/:schoolId', async (req, res) => {
  try {
    const { schoolId } = req.params;
    const formData = req.body;

    if (!schoolId || typeof formData !== 'object') {
      return res.status(400).json({ message: 'Invalid schoolId or form data' });
    }

    const sanitizedData = {
      schoolId,
      schoolName: String(formData.schoolName || 'N/A'),
      workshopFees: Number(formData.workshopFees) || 11000,
      invoiceStatus: String(formData.invoiceStatus || 'N/A'),
      kitCharges: Number(formData.kitCharges) || 0,
      kitOutDate: formData.kitOutDate ? new Date(formData.kitOutDate) : null,
      kitReceivedBy: String(formData.kitReceivedBy || 'N/A'),
      status: String(formData.status || 'N/A'),
      contentRoyalty: Number(formData.contentRoyalty) || 0,
      bookCharges: Number(formData.bookCharges) || 0,
      trainerName1: String(formData.trainerName1 || 'N/A'),
      trainerName2: String(formData.trainerName2 || 'N/A'),
      trainerRemunerationAmount: Number(formData.trainerRemunerationAmount) || 0,
      trainerTotalCost: Number(formData.trainerTotalCost) || 0,
      trainerRemunerationStatus: String(formData.trainerRemunerationStatus || 'Pending'),
      travellingAllowance: Number(formData.travellingAllowance) || 0,
      paymentDate: formData.paymentDate ? new Date(formData.paymentDate) : null,
      bltAccount: String(formData.bltAccount || 'N/A'),
      paymentMethod: String(formData.paymentMethod || 'CASH'),
      petrolDiesel: Number(formData.petrolDiesel) || 0,
      financialStatus: String(formData.financialStatus || 'Pending'),
      kitPaymentStatus: String(formData.kitPaymentStatus || 'Unpaid'),
      upiId: String(formData.upiId || 'N/A'),
      transactionId: String(formData.transactionId || 'N/A'),
      trainer1: String(formData.trainer1 || formData.trainerName1 || 'N/A'),
      trainer2: String(formData.trainer2 || formData.trainerName2 || 'N/A'),
      updatedAt: new Date()
    };

    // Calculate total expenses and remaining balance
    const totalExpenses = sanitizedData.kitCharges +
                         sanitizedData.contentRoyalty +
                         sanitizedData.bookCharges +
                         sanitizedData.travellingAllowance +
                         sanitizedData.petrolDiesel +
                         sanitizedData.trainerRemunerationAmount;
    sanitizedData.totalExpenses = totalExpenses;
    sanitizedData.total = sanitizedData.workshopFees - totalExpenses;

    // Validate numeric fields
    const numericFields = [
      'workshopFees',
      'kitCharges',
      'contentRoyalty',
      'bookCharges',
      'trainerRemunerationAmount',
      'trainerTotalCost',
      'travellingAllowance',
      'petrolDiesel',
      'totalExpenses'
    ];
    for (const field of numericFields) {
      if (isNaN(sanitizedData[field]) || sanitizedData[field] < 0) {
        return res.status(400).json({ message: `${field} must be a non-negative number` });
      }
      if (field === 'trainerRemunerationAmount' && sanitizedData[field] > 100000) {
        return res.status(400).json({ message: 'Trainer Remuneration Amount cannot exceed 100,000' });
      }
    }

    // Validate trainerTotalCost consistency
    const calculatedTrainerTotalCost = sanitizedData.trainerRemunerationAmount + sanitizedData.petrolDiesel + sanitizedData.travellingAllowance;
    if (Math.abs(sanitizedData.trainerTotalCost - calculatedTrainerTotalCost) > 0.01) {
      return res.status(400).json({ message: `Total Trainer Cost (${sanitizedData.trainerTotalCost}) does not match calculated value (${calculatedTrainerTotalCost})` });
    }

    // Validate dates
    if (sanitizedData.kitOutDate && isNaN(sanitizedData.kitOutDate.getTime())) {
      return res.status(400).json({ message: 'Invalid Kit Out Date' });
    }
    if (sanitizedData.paymentDate && isNaN(sanitizedData.paymentDate.getTime())) {
      return res.status(400).json({ message: 'Invalid Payment Date' });
    }

    // Validate select fields
    const selectFields = [
      { field: 'financialStatus', name: 'Financial Status' },
      { field: 'kitPaymentStatus', name: 'Kit Payment Status' },
      { field: 'trainerRemunerationStatus', name: 'Trainer Remuneration Status' }
    ];
    for (const { field, name } of selectFields) {
      if (!sanitizedData[field] || sanitizedData[field] === 'N/A') {
        return res.status(400).json({ message: `${name} is required` });
      }
    }

    // Log sanitized data for debugging
    console.log('Saving data to Firestore:', {
      schoolId,
      trainerRemunerationAmount: sanitizedData.trainerRemunerationAmount,
      trainerTotalCost: sanitizedData.trainerTotalCost,
      totalExpenses: sanitizedData.totalExpenses,
      total: sanitizedData.total
    });

    // Update Firestore
    await db.collection('payments').doc(schoolId).set(sanitizedData, { merge: true });
    res.status(200).json({
      message: 'Financial details updated successfully',
      data: {
        schoolId,
        trainerRemunerationAmount: sanitizedData.trainerRemunerationAmount,
        trainerTotalCost: sanitizedData.trainerTotalCost,
        totalExpenses: sanitizedData.totalExpenses,
        total: sanitizedData.total
      }
    });
  } catch (error) {
    console.error('Error updating financial details:', error);
    res.status(500).json({ message: 'Failed to update financial details', error: error.message });
  }
});
// Route to get financial details of a specific school
app.get('/financial/:schoolId', async (req, res) => {
  try {
    const { schoolId } = req.params;
    const doc = await db.collection('payments').doc(schoolId).get();
    if (!doc.exists) {
      return res.status(404).json({ message: 'School not found' });
    }
    const data = doc.data();

    // Calculate total expenses and remaining balance
    const totalExpenses = (Number(data.kitCharges) || 0) +
                         (Number(data.contentRoyalty) || 0) +
                         (Number(data.bookCharges) || 0) +
                         (Number(data.travellingAllowance) || 0) +
                         (Number(data.petrolDiesel) || 0) +
                         (Number(data.trainerRemunerationAmount) || 0);
    const total = (Number(data.workshopFees) || 11000) - totalExpenses;
    const trainerTotalCost = (Number(data.trainerRemunerationAmount) || 0) +
                            (Number(data.petrolDiesel) || 0) +
                            (Number(data.travellingAllowance) || 0);

    const responseData = {
      schoolId,
      schoolName: data.schoolName || 'N/A',
      workshopFees: Number(data.workshopFees) || 11000,
      invoiceStatus: data.invoiceStatus || 'N/A',
      kitCharges: Number(data.kitCharges) || 0,
      kitOutDate: data.kitOutDate?.toDate() ? data.kitOutDate.toDate().toISOString().split('T')[0] : 'N/A',
      kitReceivedBy: data.kitReceivedBy || 'N/A',
      status: data.status || 'N/A',
      contentRoyalty: Number(data.contentRoyalty) || 0,
      bookCharges: Number(data.bookCharges) || 0,
      trainerName1: data.trainerName1 || 'N/A',
      trainerName2: data.trainerName2 || 'N/A',
      trainerRemunerationAmount: Number(data.trainerRemunerationAmount) || 0,
      trainerTotalCost: Number(data.trainerTotalCost) || trainerTotalCost,
      trainerRemunerationStatus: data.trainerRemunerationStatus || 'Pending',
      travellingAllowance: Number(data.travellingAllowance) || 0,
      paymentDate: data.paymentDate?.toDate() ? data.paymentDate.toDate().toISOString().split('T')[0] : 'N/A',
      bltAccount: data.bltAccount || 'N/A',
      paymentMethod: data.paymentMethod || 'CASH',
      petrolDiesel: Number(data.petrolDiesel) || 0,
      totalExpenses,
      total,
      financialStatus: data.financialStatus || 'Pending',
      kitPaymentStatus: data.kitPaymentStatus || 'Unpaid',
      upiId: data.upiId || 'N/A',
      transactionId: data.transactionId || 'N/A'
    };

    // Log response data for debugging
    console.log('Returning school data:', {
      schoolId,
      trainerRemunerationAmount: responseData.trainerRemunerationAmount,
      trainerTotalCost: responseData.trainerTotalCost,
      totalExpenses: responseData.totalExpenses,
      total: responseData.total
    });

    res.status(200).json(responseData);
  } catch (error) {
    console.error('Error fetching school details:', error);
    res.status(500).json({ message: 'Failed to fetch school details', error: error.message });
  }
});

app.get('/financial/kit-out-date/:schoolId', async (req, res) => {
  try {
    const { schoolId } = req.params;
    const doc = await db.collection('schools').doc(schoolId).get();
    if (!doc.exists) {
      return res.status(404).json({ message: 'School not found' });
    }
    const data = doc.data();
    const kitOutDate = data.kitOutDate || 'N/A';
    res.status(200).json({ kitOutDate });
  } catch (error) {
    console.error('Error fetching kit out date:', error);
    res.status(500).json({ message: 'Failed to fetch kit out date' });
  }
});
//update-kit-out-date on logistic dasshboard
app.post('/update-kit-out-date', async (req, res) => {
  const { id, kitOutDate } = req.body;

  // Validate input
  if (!id || !kitOutDate) {
    return res.status(400).json({ error: 'Missing required fields: id and kitOutDate' });
  }

  try {
    // Update the kitOutDate in the schools collection
    await db.collection('schools').doc(id).update({
      kitOutDate: kitOutDate
    });

    res.status(200).json({ message: 'Kit out date updated successfully' });
  } catch (err) {
    console.error('Error updating kit out date:', err);
    if (err.code === 'not-found') {
      return res.status(404).json({ error: 'School not found' });
    }
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Update Delivery Details Route
app.post('/update-delivery', async (req, res) => {
  try {
    const { id, deliveryMode, deliveryNumber, agency } = req.body;

    // Validate inputs
    if (!id) {
      return res.status(400).send('School ID is required');
    }
    if (!['pending', 'by hand', 'post', 'train', 'bus', 'courier'].includes(deliveryMode)) {
      return res.status(400).send('Invalid delivery mode');
    }

    // Prepare update data
    const updateData = {
      deliveryMode: deliveryMode || 'pending',
      deliveryNumber: (deliveryMode === 'bus' || deliveryMode === 'train' || deliveryMode === 'post') ? (deliveryNumber || '') : '',
      agency: agency || ''
    };

    await db.collection('schools').doc(id).update(updateData);
    res.status(200).send('Delivery updated successfully');
  } catch (error) {
    console.error('Error updating delivery:', error);
    res.status(500).send('Error updating delivery');
  }
});

// Update Status Route
app.post('/update-status', async (req, res) => {
  try {
    const { id, status } = req.body;

    // Validate inputs
    if (!id) {
      return res.status(400).send('School ID is required');
    }
    if (!['pending', 'pack', 'delivered'].includes(status)) {
      return res.status(400).send('Invalid status');
    }

    // Update status in Firestore
    await db.collection('schools').doc(id).update({
      status: status,
      isCompleted: status === 'delivered' // Maintain compatibility with isCompleted
    });

    res.status(200).send('Status updated successfully');
  } catch (error) {
    console.error('Error updating status:', error);
    res.status(500).send('Error updating status');
  }
});


app.get("/admin-dashboard/logistics", async (req, res) => {
  try {
    const snapshot = await db
      .collection("schools")
      .where("isApproved", "==", true)
      .get();

    const schools = [];

    snapshot.forEach(doc => {
      const data = doc.data();
      schools.push({
        id: doc.id,
        name: data.schoolName || "N/A",
        civicsSirNumber: data.civicsTeacherNumber || "N/A",
        schoolPhoneNumber: data.schoolPhoneNumber || "N/A",
        principalNumber: data.principalNumber || "N/A",
        city: data.city || "N/A",
        district: data.district || "N/A",
        rawEventDate: data.eventDate?.toDate(), // For sorting if needed
        eventDate: data.eventDate
          ? data.eventDate.toDate().toLocaleDateString("en-IN", {
              day: "numeric",
              month: "long",
              year: "numeric",
            })
          : "Not set",
        status: data.isCompleted ? "Delivered" : "Pending",
      });
    });

    res.render("admin-logistics", { schools });
  } catch (error) {
    console.error("Error fetching logistics data:", error);
    res.status(500).send("Error loading logistics dashboard");
  }
});


//formdata & facebook routes

const FormData = require('form-data');
const fetch = require('node-fetch');

const PAGE_ACCESS_TOKEN = process.env.FB_PAGE_ACCESS_TOKEN;
const PAGE_ID = process.env.FB_PAGE_ID;

async function postPhotoToFacebook(buffer, filename, caption = '') {
  const form = new FormData();
  form.append('source', buffer, {
    filename: filename,
    contentType: 'image/jpeg'
  });
  form.append('caption', caption); // <-- Add caption here
  form.append('access_token', PAGE_ACCESS_TOKEN);

  const res = await fetch(`https://graph.facebook.com/${PAGE_ID}/photos`, {
    method: 'POST',
    body: form
  });

  const data = await res.json();
  if (!res.ok) throw new Error(data.error?.message || 'Facebook photo upload failed');
}

async function postVideoToFacebook(buffer, filename, caption = '') {
  const form = new FormData();
  form.append('source', buffer, {
    filename: filename,
    contentType: 'video/mp4'
  });
  form.append('description', caption); // <-- Corrected and placed outside the buffer object
  form.append('access_token', PAGE_ACCESS_TOKEN);

  const res = await fetch(`https://graph.facebook.com/${PAGE_ID}/videos`, {
    method: 'POST',
    body: form
  });

  const data = await res.json();
  if (!res.ok) throw new Error(data.error?.message || 'Facebook video upload failed');
}

module.exports = {
  postPhotoToFacebook,
  postVideoToFacebook
};

//instagram routes

const instagramBusinessId = process.env.INSTAGRAM_BUSINESS_ACCOUNT_ID;
const accessToken = process.env.INSTAGRAM_LONG_LIVED_ACCESS_TOKEN;

async function postToInstagram(mediaUrl, caption) {
  try {
    const mediaRes = await fetch(
      `https://graph.facebook.com/v19.0/${instagramBusinessId}/media`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: new URLSearchParams({
          image_url: mediaUrl,
          caption,
          access_token: accessToken,
        }),
      }
    );

    const mediaData = await mediaRes.json();
    if (!mediaData.id) {
      throw new Error(`Media creation failed: ${JSON.stringify(mediaData)}`);
    }

    const publishRes = await fetch(
      `https://graph.facebook.com/v19.0/${instagramBusinessId}/media_publish`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: new URLSearchParams({
          creation_id: mediaData.id,
          access_token: accessToken,
        }),
      }
    );

    const publishData = await publishRes.json();
    if (!publishData.id) {
      throw new Error(`Publish failed: ${JSON.stringify(publishData)}`);
    }

    console.log("✅ Instagram post published with ID:", publishData.id);
    return publishData;
  } catch (err) {
    console.error("❌ Failed to post to Instagram:", err.message);
    throw err;
  }
}

module.exports = { postToInstagram };



// facebook & instagram post api 

app.post('/post-to-socials', async (req, res) => {
  try {
    const { id, url, type, caption } = req.body;

    // Fetch media from URL
    const response = await fetch(url);
    const buffer = await response.buffer();

    // Post to Facebook
    if (type === 'image') {
      await postPhotoToFacebook(buffer, `${id}.jpg`, caption);
    } else if (type === 'video') {
      await postVideoToFacebook(buffer, `${id}.mp4`, caption);
    } else {
      return res.status(400).json({ message: 'Unsupported media type' });
    }

    // Post to Instagram if image
    if (type === 'image') {
      await postToInstagram(url, caption || '');
    }

    res.status(200).json({ message: 'Posted to Facebook and Instagram successfully' });
  } catch (error) {
    // console.error('❌ Failed to post to social platforms:', error);
    res.status(500).json({ message: 'Error posting to social platforms', error: error.message });
  }
});

// Update description
app.put('/update-description/:id', async (req, res) => {
  const { id } = req.params;
  const { description } = req.body;

  try {
    const mediaRef = db.collection('mediaUploads').doc(id);
    await mediaRef.update({ description });

    res.status(200).json({ message: 'Description updated successfully' });
  } catch (err) {
    console.error('Error updating description:', err);
    res.status(500).json({ message: 'Internal Server Error' });
  }
});
// Register Coordinator Route
// Middleware to check if coordinator is authenticated
const isAuthenticated = (req, res, next) => {
    if (!req.session.coordinator || req.session.coordinator.role !== 'coordinator') {
        // If it's an AJAX or API request
        if (req.xhr || req.headers.accept.includes('application/json')) {
            return res.status(401).json({ error: 'Not authenticated' });
        }
        // Otherwise, it's a normal browser request
        return res.redirect('/login-coordinator');
    }
    next();
};


app.post("/upload/pdf", uploadpdf, async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No PDF uploaded" });

    const file = req.file;
    const fileName = `pdfs/${Date.now()}_${file.originalname}`;
    const fileRef = bucket.file(fileName);

    await fileRef.save(file.buffer, {
      contentType: file.mimetype,
      public: true, // make accessible via URL
    });

    const pdfUrl = `https://storage.googleapis.com/${bucket.name}/${fileName}`;

    // Save metadata to Firestore
    await db.collection("pdfUploads").add({
      name: file.originalname,
      url: pdfUrl,
      uploadedAt: new Date(),
    });

    res.json({ message: "PDF uploaded successfully", url: pdfUrl });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Upload failed" });
  }
});


// Get all uploaded PDFs
app.get("/get/pdfs", async (req, res) => {
  try {
    const snapshot = await db.collection("pdfUploads").orderBy("uploadedAt", "desc").get();

    const pdfs = snapshot.docs.map(doc => ({
      id: doc.id,
      name: doc.data().name,
      url: doc.data().url,
      uploadedAt: doc.data().uploadedAt?.toDate ? doc.data().uploadedAt.toDate() : null,
    }));

    res.json({ pdfs });
  } catch (err) {
    console.error("Error fetching PDFs:", err);
    res.status(500).json({ error: "Failed to fetch PDFs" });
  }
});


    // Register Coordinator Route
  app.post('/co-ordinator-register', async (req, res) => {
    try {
        const { name, number, email, organization, password } = req.body;
        if (!name || !number || !email || !organization || !password) {
            return res.status(400).render('co-ordinator-register', { error: 'All fields are required', success: null });
        }

        await db.collection('coordinators').add({
            name: name.trim(),
            number: number.trim(),
            email: email.trim(),
            organization: organization.trim(),
            password: password.trim(),
            createdAt: admin.firestore.FieldValue.serverTimestamp()
        });

        res.redirect('/admin-dashboard?success=' + encodeURIComponent('Coordinator registered successfully'));
    } catch (error) {
        console.error('Registration error:', error);
        res.redirect('/register-page?error=' + encodeURIComponent('Something went wrong: ' + error.message));
    }
});

    // Login Page Route
    app.get('/login-coordinator', (req, res) => {
        res.render('co-ordinator', { error: null, success: null });
    });

app.post('/login-coordinator', async (req, res) => {
    try {
        console.log("Incoming login body:", req.body);

        const number = String(req.body.number || '').trim();
        const password = String(req.body.password || '').trim();

        if (!number || !password) {
            return res.status(400).render('co-ordinator', {
                error: 'Phone number and password are required',
                success: null
            });
        }

        // Find coordinator by number
        const snapshot = await db.collection('coordinators')
            .where('number', '==', number)
            .limit(1)
            .get();

        if (snapshot.empty) {
            return res.status(401).render('co-ordinator', {
                error: 'Invalid phone number or password',
                success: null
            });
        }

        const coordinator = snapshot.docs[0].data();

        // Ensure password field exists
        if (!coordinator.password) {
            console.error("Coordinator record has no password field:", coordinator);
            return res.status(500).render('co-ordinator', {
                error: 'Your account is missing a password. Contact admin.',
                success: null
            });
        }

        // Plain text password comparison
        if (String(coordinator.password).trim() !== password) {
            return res.status(401).render('co-ordinator', {
                error: 'Invalid phone number or password',
                success: null
            });
        }

        // Save session
        req.session.coordinator = {
            id: snapshot.docs[0].id,
            name: coordinator.name,
            number: coordinator.number,
            organization: coordinator.organization,
            role: 'coordinator'
        };

        req.session.save(() => {
            res.redirect('/coordinator-dashboard');
        });

    } catch (error) {
        console.error('Error logging in:', error.message, error.stack);
        res.status(500).render('co-ordinator', {
            error: 'Something went wrong. Please try again.',
            success: null
        });
    }
});



// GET: Coordinator Dashboard
app.get('/coordinator-dashboard', isAuthenticated, async (req, res) => {
    try {
        // Validate session
        if (!req.session.coordinator?.id) {
            console.error('Invalid session: Coordinator ID missing');
            return res.status(401).redirect('/login?error=' + encodeURIComponent('Session expired. Please log in again.'));
        }

        const coordinatorId = req.session.coordinator.id;
        const trainerId = req.query.trainerId || null;

        // Fetch approved schools
        const approvedSnapshot = await db
            .collection('schools')
            .where('isApproved', '==', true)
            .where('coordinatorId', '==', coordinatorId)
            .get();

        const schools = approvedSnapshot.docs
            .map(doc => {
                const data = doc.data();
                const eventDate = data.eventDate?.toDate
                    ? data.eventDate.toDate()
                    : data.eventDate
                        ? new Date(data.eventDate)
                        : null;

                return {
                    id: doc.id,
                    name: data.schoolName || 'N/A',
                    civicsSirNumber: data.civicsTeacherNumber || 'N/A',
                    schoolPhoneNumber: data.schoolPhoneNumber || 'N/A',
                    principalNumber: data.principalNumber || 'N/A',
                    city: data.city || 'N/A',
                    district: data.district || 'N/A',
                    eventDate: eventDate ? eventDate.toLocaleDateString('en-IN') : 'Not set',
                    status: data.coordinatorStatus || 'inprogress',
                    trainerId1: data.trainerId1 || null,
                    trainerId2: data.trainerId2 || null,
                    workshopStartTime: data.workshopStartTime || null,
                    workshopEndTime: data.workshopEndTime || null,
                    resourcesConfirmed: data.resourcesConfirmed || false,
                    coordinatorId: data.coordinatorId || null,
                    agency: data.agency || '',
                    deliveryMode: data.deliveryMode || 'pending',
                    registerStatus: data.registerStatus || 'Yes',
                    mcqStatus: data.mcqStatus || 'Pending',
                    approvalStatus: data.mcqStatus === 'Completed' ? 'Approved' : 'Approved, MCQ Remaining'
                };
            })
            .sort((a, b) => {
                if (!a.eventDate || a.eventDate === 'Not set') return 1;
                if (!b.eventDate || b.eventDate === 'Not set') return -1;
                return new Date(a.eventDate) - new Date(b.eventDate);
            });

        // Fetch visited schools
        let visitedQuery = db
            .collection('visitedSchools')
            .where('coordinatorId', '==', coordinatorId);
        if (trainerId) {
            visitedQuery = visitedQuery.where('trainerId', '==', trainerId);
        }
        const visitedSnapshot = await visitedQuery.get();

        const visitedSchools = visitedSnapshot.docs.map(doc => {
            const data = doc.data();
            const visitDate = data.visitDate
                ? new Date(data.visitDate).toLocaleDateString('en-IN')
                : 'Not set';
            return {
                id: doc.id,
                name: data.name || 'N/A',
                visitDate: visitDate,
                schoolAddress: data.schoolAddress || 'N/A',
                contactPerson: data.contactPerson || 'N/A',
                contactNumber: data.contactNumber || 'N/A',
                visitNotes: data.visitNotes || 'N/A',
                coordinatorId: data.coordinatorId || null
            };
        });

        // Fetch completed schools
        let completedQuery = db
            .collection('schools')
            .where('isCompleted', '==', true)
            .where('coordinatorId', '==', coordinatorId);
        if (trainerId) {
            completedQuery = completedQuery.where('trainerId1', '==', trainerId);
        }
        const completedSnapshot = await completedQuery.get();

        const completedSchools = completedSnapshot.docs.map(doc => {
            const data = doc.data();
            const eventDate = data.eventDate?.toDate
                ? data.eventDate.toDate()
                : data.eventDate
                    ? new Date(data.eventDate)
                    : null;
            return {
                id: doc.id,
                name: data.schoolName || 'N/A',
                principalName: data.principalName || data.principalNumber || 'N/A',
                schoolPhoneNumber: data.schoolPhoneNumber || 'N/A',
                email: data.email || 'N/A',
                contactPersonNumber: data.contactPersonNumber || data.principalNumber || 'N/A',
                eventDate: eventDate ? eventDate.toLocaleDateString('en-IN') : 'Not set',
                trainerId1: data.trainerId1 || null,
                trainerId2: data.trainerId2 || null,
                workshopStartTime: data.workshopStartTime || null,
                workshopEndTime: data.workshopEndTime || null,
                resourcesConfirmed: data.resourcesConfirmed || false,
                coordinatorId: data.coordinatorId || null,
                agency: data.agency || '',
                deliveryMode: data.deliveryMode || 'pending',
                status: data.coordinatorStatus || 'inprogress',
                registerStatus: 'Yes',
                mcqStatus: data.mcqStatus || 'Completed',
                approvalStatus: 'Approved'
            };
        });

        // Fetch trainers
const trainersSnapshot = await db.collection('trainers').get();
const trainers = trainersSnapshot.docs.map(doc => ({
    id: doc.id,
    name: doc.data().trainerName || 'Unknown Trainer',
    mobileNumber: doc.data().mobileNumber || 'N/A' 
}));

        // Fetch feedback data
        const feedbackSnapshot = await db
            .collection('studentFeedback')
            .where('coordinatorId', '==', coordinatorId)
            .get();

        const feedbackData = feedbackSnapshot.docs.map(doc => {
            const data = doc.data();
            return {
                id: doc.id,
                studentName: data.studentName || 'N/A',
                className: data.className || 'N/A',
                question1: data.question1 || 'N/A',
                question2: data.question2 || 'N/A',
                question3: data.question3 || 'N/A',
                question4: data.question4 || 'N/A',
                question5: data.question5 || 'N/A',
                status: data.status || 'submitted',
                createdAt: data.createdAt?.toDate
                    ? data.createdAt.toDate().toLocaleDateString('en-IN')
                    : 'N/A',
                updatedAt: data.updatedAt?.toDate
                    ? data.updatedAt.toDate().toLocaleDateString('en-IN')
                    : 'N/A'
            };
        });

        // Fetch workshop summaries
        const workshopSummariesSnapshot = await db
            .collection('workshopSummaries')
            .where('coordinatorId', '==', coordinatorId)
            .get();

        const workshopSummaries = workshopSummariesSnapshot.docs.map(doc => {
            const data = doc.data();
            const workshopDate = data.workshopDate
                ? new Date(data.workshopDate).toLocaleDateString('en-IN')
                : 'Not set';
            const coordinatorDate = data.coordinatorDate
                ? new Date(data.coordinatorDate).toLocaleDateString('en-IN')
                : 'Not set';
            return {
                id: doc.id,
                schoolName: data.schoolName || 'N/A',
                schoolAddress: data.schoolAddress || 'N/A',
                workshopDate: workshopDate,
                trainer1: data.trainer1 || 'N/A',
                trainer2: data.trainer2 || 'N/A',
                coordinatorName: data.coordinatorName || 'N/A',
                techSupport: data.techSupport || 'N/A',
                principalName: data.principalName || 'N/A',
                financialStatus: data.financialStatus || 'N/A',
                kitPaymentStatus: data.kitPaymentStatus || 'N/A',
                trainerRemunerationStatus: data.trainerRemunerationStatus || 'N/A',
                paymentMode: data.paymentMode || 'N/A',
                transactionId: data.transactionId || 'N/A',
                coordinatorDeclaration: data.coordinatorDeclaration ? 'Yes' : 'No',
                coordinatorDate: coordinatorDate,
                coordinatorPlace: data.coordinatorPlace || 'N/A',
                createdAt: data.createdAt?.toDate
                    ? data.createdAt.toDate().toLocaleDateString('en-IN')
                    : 'N/A'
            };
        });

        res.render('coordinator-dashboard', {
            coordinator: req.session.coordinator,
            schools,
            visitedSchools,
            completedSchools,
            trainers,
            trainerId,
            feedbackData,
            workshopSummaries,
            error: req.query.error || null,
            success: req.query.success || null
        });
    } catch (error) {
        console.error('Error fetching dashboard data:', error.message, error.stack);
        res.status(500).render('coordinator-dashboard', {
            coordinator: req.session.coordinator || {},
            schools: [],
            visitedSchools: [],
            completedSchools: [],
            trainers: [],
            trainerId: null,
            feedbackData: [],
            workshopSummaries: [],
            error: 'Failed to load dashboard data',
            success: null
        });
    }
});

// Save visited school to Firestore
app.post('/submit-visited-school', isAuthenticated, async (req, res) => {
    try {
        // Check if req.body is undefined or empty
        if (!req.body || Object.keys(req.body).length === 0) {
            return res.status(400).json({
                success: false,
                error: 'Request body is empty or undefined'
            });
        }

        const {
            visitedSchoolName,
            visitDate,
            schoolAddress,
            contactPerson,
            contactNumber,
            visitNotes
        } = req.body;

        // Validate required fields
        if (!visitedSchoolName || !visitDate || !schoolAddress) {
            return res.status(400).json({
                success: false,
                error: 'Missing required fields: visitedSchoolName, visitDate, and schoolAddress are required'
            });
        }

        // Add coordinatorId to the document
        const result = await db.collection('visitedSchools').add({
            name: visitedSchoolName,
            visitDate,
            schoolAddress,
            contactPerson: contactPerson || '',
            contactNumber: contactNumber || '',
            visitNotes: visitNotes || '',
            coordinatorId: req.session.coordinator?.id || null,
            createdAt: new Date()
        });

        console.log(`✅ Visited school added: ${visitedSchoolName}, ID: ${result.id}`);
        return res.status(200).json({
            success: true,
            message: 'School visit recorded successfully',
            redirect: '/coordinator-dashboard?success=' + encodeURIComponent('School visit added successfully')
        });
    } catch (error) {
        console.error('❌ Error adding visited school:', error.message, error.stack);
        return res.status(500).json({
            success: false,
            error: 'Failed to add school visit',
            details: error.message
        });
    }
});

app.get('/visited-schools', isAuthenticated, async (req, res) => {
    try {
        const coordinatorId = req.session.coordinator?.id;
        if (!coordinatorId) {
            return res.status(401).json({
                success: false,
                error: 'Unauthorized: Coordinator ID missing'
            });
        }

        const snapshot = await db
            .collection('visitedSchools')
            .where('coordinatorId', '==', coordinatorId)
            .get();

        const visitedSchools = snapshot.docs.map(doc => ({
            id: doc.id,
            name: doc.data().name || 'N/A',
            visitDate: doc.data().visitDate ? new Date(doc.data().visitDate).toLocaleDateString('en-IN') : 'Not set',
            schoolAddress: doc.data().schoolAddress || 'N/A',
            contactPerson: doc.data().contactPerson || 'N/A',
            contactNumber: doc.data().contactNumber || 'N/A',
            visitNotes: doc.data().visitNotes || 'N/A'
        }));

        return res.status(200).json({
            success: true,
            visitedSchools
        });
    } catch (error) {
        console.error('❌ Error fetching visited schools:', error.message, error.stack);
        return res.status(500).json({
            success: false,
            error: 'Failed to fetch visited schools',
            details: error.message
        });
    }
});

//feedback summary formr for student
app.get('/feedback', async (req, res) => {
    try {
        const filters = {
            coordinatorId: req.query.coordinatorId // Assuming coordinatorId is passed as query param
        };

        let feedbackEntries = [];
        let feedbackQuery = db.collection('studentFeedback');
        
        if (filters.coordinatorId) {
            feedbackQuery = feedbackQuery.where('coordinatorId', '==', filters.coordinatorId);
        }
        
        const feedbackSnapshot = await feedbackQuery
            .where('status', '==', 'submitted')
            .orderBy('createdAt', 'desc')
            .get();
            
        feedbackEntries = feedbackSnapshot.docs.map(doc => {
            const data = doc.data();
            return {
                id: doc.id,
                studentName: data.studentName || 'N/A',
                className: data.className || 'N/A',
                question1: data.question1 || 'N/A',
                question2: data.question2 || 'N/A',
                question3: data.question3 || 'N/A',
                question4: data.question4 || 'N/A',
                question5: data.question5 || 'N/A',
                coordinatorId: data.coordinatorId || 'N/A',
                createdAt: formatToIST(data.createdAt),
                status: data.status || 'N/A'
            };
        });

        // Render the EJS template with feedback entries
        res.render('feedback', { feedbackEntries });

    } catch (error) {
        console.error('Error fetching student feedback:', error);
        res.status(500).send('Error fetching feedback');
    }
});
 // GET: Download Feedback Excel Template
app.get('/download-feedback-template', (req, res) => {
    try {
        const workbook = xlsx.utils.book_new();
        const worksheetData = [['Student Name', 'Class', 'Question 1', 'Question 2', 'Question 3', 'Question 4', 'Question 5']];
        const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Student Feedback');

        const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        res.setHeader('Content-Disposition', 'attachment; filename=student_feedback_template.xlsx');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (error) {
        console.error('Error generating feedback template:', error.message, error.stack);
        res.status(500).redirect('/coordinator-dashboard?error=' + encodeURIComponent('Error generating feedback template'));
    }
});

    // POST: Upload Student Feedback Excel File
app.post('/upload-feedback-excel', isAuthenticated, uploadFeedback, async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        // Parse Excel file
        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet, { defval: '' });

        if (data.length === 0) {
            return res.status(400).json({ error: 'Excel file is empty' });
        }

        // Validate headers
        const expectedHeaders = ['Student Name', 'Class', 'Question 1', 'Question 2', 'Question 3', 'Question 4', 'Question 5'];
        const headers = Object.keys(data[0] || {}).map(h => h.trim());
        const missingHeaders = expectedHeaders.filter(h => !headers.includes(h));
        if (missingHeaders.length > 0) {
            return res.status(400).json({ error: `Missing required headers: ${missingHeaders.join(', ')}` });
        }

        // Validate data rows
        const feedbackEntries = data.map((row, index) => {
            const studentName = row['Student Name']?.toString().trim();
            const className = row['Class']?.toString().trim();
            const question1 = row['Question 1']?.toString().trim();
            const question2 = row['Question 2']?.toString().trim();
            const question3 = row['Question 3']?.toString().trim();
            const question4 = row['Question 4']?.toString().trim();
            const question5 = row['Question 5']?.toString().trim();

            // Required fields validation
            if (!studentName || !className || !question1 || !question2 || !question3 || !question4 || !question5) {
                throw new Error(`Row ${index + 2}: All fields (Student Name, Class, Question 1-5) are required`);
            }

            return {
                studentName,
                className,
                question1,
                question2,
                question3,
                question4,
                question5,
                status: 'submitted',
                coordinatorId: req.session.coordinator.id,
                createdAt: admin.firestore.FieldValue.serverTimestamp(),
                updatedAt: admin.firestore.FieldValue.serverTimestamp()
            };
        });

        // Clear existing feedback (optional, retained from original)
        const feedbackSnapshot = await db.collection('studentFeedback')
            .where('status', '==', 'submitted')
            .where('coordinatorId', '==', req.session.coordinator.id)
            .get();
        const deleteBatch = db.batch();
        feedbackSnapshot.docs.forEach(doc => deleteBatch.delete(doc.ref));
        await deleteBatch.commit();

        // Write new feedback entries
        const writeBatch = db.batch();
        for (const feedback of feedbackEntries) {
            const feedbackRef = db.collection('studentFeedback').doc();
            writeBatch.set(feedbackRef, feedback);
        }
        await writeBatch.commit();

        res.json({ message: 'Feedback Excel file uploaded successfully! Student Feedback collection updated.' });
    } catch (error) {
        console.error('Error uploading Feedback Excel file:', error.message, error.stack);
        res.status(500).json({ error: `Failed to upload Feedback Excel file: ${error.message}` });
    }
});

// Route to fetch school tracker data, including trainer financials
app.get('/school-tracker', async (req, res) => {
  try {
    const schoolSnapshot = await db.collection('schools').get();
    const paymentSnapshot = await db.collection('payments').get();
    const participantSnapshot = await db.collection('participants').get();
    const coordinatorSnapshot = await db.collection('coordinators').get();
    const workshopSummarySnapshot = await db.collection('workshopSummaries').get();

    const schools = [];
    const payments = {};
    const participants = [];
    const coordinators = {};
    const workshopSummaries = {};

    // Process coordinators
    coordinatorSnapshot.forEach(doc => {
      coordinators[doc.id] = doc.data().name || 'Not Assigned';
    });

    // Process workshop summaries
    workshopSummarySnapshot.forEach(doc => {
      const data = doc.data();
      const key = `${data.schoolName}_${data.coordinatorId}`;
      workshopSummaries[key] = {
        civicTeachersFeedback: data.civicTeachersFeedback || 'Not Assigned',
        principalFeedback: data.principalFeedback || 'Not Assigned',
        socialMediaFeedback: data.socialMediaFeedback || 'Not Assigned',
        trainer1: data.trainer1 || 'Not Assigned',
        trainer2: data.trainer2 || 'Not Assigned',
      };
    });

    // Process payments
    paymentSnapshot.forEach(doc => {
      const data = doc.data();
      payments[doc.id] = {
        workshopFees: Number(data.workshopFees) || 0,
        invoiceStatus: data.invoiceStatus || 'Not Assigned',
        kitCharges: Number(data.kitCharges) || 0,
        kitOutDate: data.kitOutDate?.toDate?.()?.toLocaleDateString('en-IN') || 'Not Assigned',
        kitReceivedBy: data.kitReceivedBy || 'Not Assigned',
        status: data.status || 'Not Assigned',
        contentRoyalty: Number(data.contentRoyalty) || 0,
        bookCharges: Number(data.bookCharges) || 0,
        trainerName1: data.trainerName1 || 'Not Assigned',
        trainerName2: data.trainerName2 || 'Not Assigned',
        trainerRemuneration: Number(data.trainerRemunerationAmount || data.trainerHonorarium) || 0,
        travellingAllowance: Number(data.travellingAllowance) || 0,
        petrolDiesel: Number(data.petrolDiesel) || 0,
        trainerRemunerationStatus: data.trainerRemunerationStatus || 'Pending',
        paid: Number(data.paid) || 0,
        paymentDate: data.paymentDate?.toDate?.()?.toLocaleDateString('en-IN') || 'Not Assigned',
        bltAccount: Number(data.bltAccount) || 0, 
        paymentMethod: data.paymentMethod || 'CASH',
        total: Number(data.total) || 0,
        trainerTotalCost: Number(data.trainerTotalCost) || 0,
        totalExpenses: Number(data.totalExpenses) || 0,
      };
    });

    // Process participants
    participantSnapshot.forEach(doc => {
      const data = doc.data();
      participants.push({
        schoolNameDropdown: data.schoolNameDropdown || 'Not Assigned',
        hasCompletedMCQ: data.hasCompletedMCQ || false,
        isChampion: data.isChampion || false,
        feedback: data.feedback || '',
      });
    });

    // Process schools and merge data
    schoolSnapshot.forEach(doc => {
      const data = doc.data();
      const schoolId = doc.id;
      const paymentData = payments[schoolId] || {};
      const coordinatorName = coordinators[data.coordinatorId] || 'Not Assigned';
      const feedbackKey = `${data.schoolName}_${data.coordinatorId}`;
      const feedbackData = workshopSummaries[feedbackKey] || {
        civicTeachersFeedback: 'Not Assigned',
        principalFeedback: 'Not Assigned',
        socialMediaFeedback: 'Not Assigned',
        trainer1: 'Not Assigned',
        trainer2: 'Not Assigned',
      };

      let eventDate = data.eventDate?.toDate?.() || (typeof data.eventDate === 'string' ? new Date(data.eventDate) : null);

      const trainerName = [
        feedbackData.trainer1,
        feedbackData.trainer2,
      ].filter(t => t && t !== 'Not Assigned').join(', ') || [
        paymentData.trainerName1,
        paymentData.trainerName2,
      ].filter(t => t && t !== 'Not Assigned').join(', ') || 'Not Assigned';

      const champFeedbackStatus = participants.some(p =>
        p.schoolNameDropdown === data.schoolName && p.feedback && p.feedback !== ''
      ) ? 'Y' : 'N';

      const thinkingLetter = data.thinkingLetter === true ? 'Yes' : data.thinkingLetter === false ? 'No' : 'Not Assigned';
      const cotation = data.cotation === true ? 'Yes' : data.cotation === false ? 'No' : 'Not Assigned';

      const totalExpenses = (paymentData.kitCharges || 0) +
                           (paymentData.contentRoyalty || 0) +
                           (paymentData.bookCharges || 0) +
                           (paymentData.travellingAllowance || 0) +
                           (paymentData.petrolDiesel || 0) +
                           (paymentData.trainerRemuneration || 0);

      const total = (paymentData.workshopFees || 0) - totalExpenses;

      schools.push({
        schoolId,
        schoolName: data.schoolName || 'Not Assigned',
        city: data.city || 'Not Assigned',
        district: data.district || 'Not Assigned',
        eventDate: eventDate?.toLocaleDateString('en-IN') || 'Not Assigned',
        isApproved: data.isApproved || false,
        deliveryMode: data.deliveryMode || 'pending',
        coordinatorName,
        workshopFees: paymentData.workshopFees || 0,
        invoiceStatus: paymentData.invoiceStatus || 'Not Assigned',
        kitCharges: paymentData.kitCharges || 0,
        kitDispatchedDate: paymentData.kitOutDate || 'Not Assigned',
        kitReceivedBy: paymentData.kitReceivedBy || 'Not Assigned',
        status: paymentData.status || 'Not Assigned',
        contentRoyalty: paymentData.contentRoyalty || 0,
        bookCharges: paymentData.bookCharges || 0,
        trainerName,
        trainerRemuneration: paymentData.trainerRemuneration || 0,
        travellingAllowance: paymentData.travellingAllowance || 0,
        petrolDiesel: paymentData.petrolDiesel || 0,
        totalTrainerCost: paymentData.trainerTotalCost || 0,
        trainerRemunerationStatus: paymentData.trainerRemunerationStatus || 'Pending',
        paymentDate: paymentData.paymentDate || 'Not Assigned',
        bltAccount: paymentData.bltAccount || 0,
        paymentMethod: paymentData.paymentMethod || 'CASH',
        total: total || 0,
        totalExpenses: totalExpenses || 0,
        civicTeachersFeedback: feedbackData.civicTeachersFeedback,
        principalFeedback: feedbackData.principalFeedback !== 'Not Assigned' && feedbackData.principalFeedback ? 'Y' : 'N',
        feedbackSocialMedia: feedbackData.socialMediaFeedback,
        centreCoordinationExpenses: paymentData.petrolDiesel || 0,
        champFeedbackStatus,
        thinkingLetter,
        cotation,
      });
    });

    const totalTrainerCost = schools.reduce((sum, school) => sum + (school.totalTrainerCost || 0), 0);
    const totalAllExpenses = schools.reduce((sum, school) => sum + (school.totalExpenses || 0), 0);

    res.render('tracker', {
      schools,
      participants,
      error: null,
      totalTrainerCost: totalTrainerCost.toLocaleString('en-IN'),
      totalAllExpenses: totalAllExpenses.toLocaleString('en-IN'),
    });

  } catch (err) {
    console.error('Error fetching data:', err);
    res.render('tracker', {
      schools: [],
      participants: [],
      error: err.message,
      totalTrainerCost: 0,
      totalAllExpenses: 0,
    });
  }
});



// Excel export route
app.get('/school-tracker/export', async (req, res) => {
  try {
    const schoolSnapshot = await db.collection('schools').get();
    const paymentSnapshot = await db.collection('payments').get();
    const participantSnapshot = await db.collection('participants').get();
    const coordinatorSnapshot = await db.collection('coordinators').get();
    const workshopSummarySnapshot = await db.collection('workshopSummaries').get();

    const schools = [];
    const payments = {};
    const participants = [];
    const coordinators = {};
    const workshopSummaries = {};

    coordinatorSnapshot.forEach(doc => {
      coordinators[doc.id] = doc.data().name || 'N/A';
    });

    workshopSummarySnapshot.forEach(doc => {
      const data = doc.data();
      const key = `${data.schoolName}_${data.coordinatorId}`;
      workshopSummaries[key] = {
        civicTeachersFeedback: data.civicTeachersFeedback || 'N/A',
        principalFeedback: data.principalFeedback || 'N/A',
        socialMediaFeedback: data.socialMediaFeedback || 'N/A',
        trainer1: data.trainer1 || 'N/A',
        trainer2: data.trainer2 || 'N/A',
      };
    });

    paymentSnapshot.forEach(doc => {
      const data = doc.data();
      payments[doc.id] = {
        workshopFees: Number(data.workshopFees) || 11000,
        invoiceStatus: data.invoiceStatus || 'N/A',
        kitCharges: Number(data.kitCharges) || 0,
        kitOutDate: data.kitOutDate?.toDate?.()?.toLocaleDateString('en-IN') || 'N/A',
        kitReceivedBy: data.kitReceivedBy || 'N/A',
        contentRoyalty: Number(data.contentRoyalty) || 0,
        bookCharges: Number(data.bookCharges) || 0,
        trainerName1: data.trainerName1 || 'N/A',
        trainerName2: data.trainerName2 || 'N/A',
        trainerRemuneration: Number(data.trainerRemunerationAmount || data.trainerHonorarium) || 0,
        travellingAllowance: Number(data.travellingAllowance) || 0,
        petrolDiesel: Number(data.petrolDiesel) || 0,
        paymentDate: data.paymentDate?.toDate?.()?.toLocaleDateString('en-IN') || 'N/A',
        bltAccount: Number(data.bltAccount) || 0, // default 0
        paymentMethod: data.paymentMethod || 'CASH',
        totalExpenses: Number(data.totalExpenses) || 0,
      };
    });

    participantSnapshot.forEach(doc => {
      const data = doc.data();
      participants.push({
        schoolNameDropdown: data.schoolNameDropdown || 'N/A',
        hasCompletedMCQ: data.hasCompletedMCQ || false,
        isChampion: data.isChampion || false,
        feedback: data.feedback || '',
      });
    });

    schoolSnapshot.forEach(doc => {
      const data = doc.data();
      const schoolId = doc.id;
      const paymentData = payments[schoolId] || {};
      const coordinatorName = coordinators[data.coordinatorId] || 'N/A';
      const feedbackKey = `${data.schoolName}_${data.coordinatorId}`;
      const feedbackData = workshopSummaries[feedbackKey] || {
        civicTeachersFeedback: 'N/A',
        principalFeedback: 'N/A',
        socialMediaFeedback: 'N/A',
        trainer1: 'N/A',
        trainer2: 'N/A',
      };

      let eventDate = data.eventDate?.toDate?.() || (typeof data.eventDate === 'string' ? new Date(data.eventDate) : null);

      const trainerName = [
        feedbackData.trainer1,
        feedbackData.trainer2,
      ].filter(t => t && t !== 'N/A').join(', ') || [
        paymentData.trainerName1,
        paymentData.trainerName2,
      ].filter(t => t && t !== 'N/A').join(', ') || 'N/A';

      const champFeedbackStatus = participants.some(p =>
        p.schoolNameDropdown === data.schoolName && p.feedback && p.feedback !== ''
      ) ? 'Y' : 'N';

      const thinkingLetter = data.thinkingLetter === true ? 'Yes' : data.thinkingLetter === false ? 'No' : 'N/A';
      const cotation = data.cotation === true ? 'Yes' : data.cotation === false ? 'No' : 'N/A';

      schools.push({
        schoolId,
        schoolName: data.schoolName || 'N/A',
        city: data.city || 'N/A',
        district: data.district || 'N/A',
        eventDate: eventDate?.toLocaleDateString('en-IN') || 'N/A',
        coordinatorName,
        workshopFees: paymentData.workshopFees || 11000,
        kitCharges: paymentData.kitCharges || 0,
        kitDispatchedDate: paymentData.kitOutDate || 'N/A',
        kitReceivedBy: paymentData.kitReceivedBy || 'N/A',
        contentRoyalty: paymentData.contentRoyalty || 0,
        bookCharges: paymentData.bookCharges || 0,
        trainerName,
        trainerRemuneration: paymentData.trainerRemuneration || 0,
        travellingAllowance: paymentData.travellingAllowance || 0,
        petrolDiesel: paymentData.petrolDiesel || 0,
        totalExpenses: paymentData.totalExpenses || 0,
        civicTeachersFeedback: feedbackData.civicTeachersFeedback,
        principalFeedback: feedbackData.principalFeedback !== 'N/A' && feedbackData.principalFeedback ? 'Y' : 'N',
        feedbackSocialMedia: feedbackData.socialMediaFeedback,
        champFeedbackStatus,
        thinkingLetter,
        cotation,
        bltAccount: paymentData.bltAccount != null ? paymentData.bltAccount : 0, // default 0
      });
    });

    const excelData = schools.map((school, index) => ({
      'S. No': index + 1,
      'School Name': school.schoolName,
      'City': school.city,
      'District': school.district,
      'Event Date': school.eventDate,
      'Event Coordinator': school.coordinatorName,
      'Student Login': participants.filter(p => p.schoolNameDropdown === school.schoolName).length,
      'MCQ Completed': participants.filter(p => p.schoolNameDropdown === school.schoolName && p.hasCompletedMCQ).length,
      'Champ Status': participants.filter(p => p.schoolNameDropdown === school.schoolName && p.isChampion).length,
      'Champ Feedback Status': school.champFeedbackStatus,
      'Feedback for Social Media': school.feedbackSocialMedia,
      'Principal Feedback (Y/N)': school.principalFeedback,
      'Civic Teachers Feedback': school.civicTeachersFeedback,
      'Workshop Fees': school.workshopFees,
      'Delivery Mode': school.deliveryMode || 'pending',
      'Kit Dispatched Date': school.kitDispatchedDate,
      'Kit Received By': school.kitReceivedBy,
      'Kit Charges': school.kitCharges,
      'Content Royalty': school.contentRoyalty,
      'Book Charges': school.bookCharges,
      'Centre Coordination Expenses': school.petrolDiesel,
      'Trainer Name': school.trainerName,
      'Trainer Remuneration': school.trainerRemuneration,
      'Travelling Allowance': school.travellingAllowance,
      'Total Expenses': school.totalExpenses,
      'BLT Account': school.bltAccount, // default 0
      'Thanking Letter': school.thinkingLetter,
      'Cotation': school.cotation,
    }));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(excelData);
    XLSX.utils.book_append_sheet(wb, ws, 'School Tracker');

    const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
    res.setHeader('Content-Disposition', 'attachment; filename=School_Tracker.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);

  } catch (err) {
    console.error('Error generating Excel:', err);
    res.status(500).json({ error: 'Failed to generate Excel file' });
  }
});



// POST endpoint for submitting workshop summary

app.post('/submit-workshop-summary', isAuthenticated, async (req, res) => {
    try {
        console.log('Received form data:', req.body);

        if (!req.body) {
            return res.status(400).json({ error: 'No form data provided' });
        }

        // Prepare form data
        const formData = {
            schoolName: req.body.schoolName?.trim(),
            schoolAddress: req.body.schoolAddress?.trim(),
            workshopDate: req.body.workshopDate?.trim(),
            studentFeedbackFormNumber: parseInt(req.body.studentFeedbackFormNumber, 50),
            trainer1: req.body.trainer1?.trim(),
            trainer2: req.body.trainer2?.trim() || '',
            coordinatorName: req.body.coordinatorName?.trim(),
            techSupport: req.body.techSupport?.trim() || '',
            principalName: req.body.principalName?.trim(),
            trainerFeedback: req.body.trainerFeedback?.trim() || '',
            socialMediaFeedback: req.body.socialMediaFeedback?.trim() || '',
            principalFeedback: req.body.principalFeedback?.trim(),
            civicTeachersFeedback: req.body.civicTeachersFeedback?.trim(),
            financialStatus: req.body.financialStatus?.trim(),
            kitPaymentStatus: req.body.kitPaymentStatus?.trim(),
            trainerRemunerationStatus: req.body.trainerRemunerationStatus?.trim(),
            paymentMode: req.body.paymentMode?.trim(),
            upiId: req.body.upiId?.trim() || '',
            transactionId: req.body.transactionId?.trim() || '',
            coordinatorDeclaration: req.body.coordinatorDeclaration === true || req.body.coordinatorDeclaration === 'true',
            coordinatorDate: req.body.coordinatorDate?.trim(),
            coordinatorPlace: req.body.coordinatorPlace?.trim(),
            coordinatorId: req.session.coordinator.id,
            createdAt: admin.firestore.FieldValue.serverTimestamp(),
            updatedAt: admin.firestore.FieldValue.serverTimestamp()
        };

        // Validate form data
        const validationErrors = validateFormData(formData);
        if (validationErrors.length > 0) {
            return res.status(400).json({ error: 'Validation failed', details: validationErrors });
        }

        // Save to Firestore
        const docRef = await db.collection('workshopSummaries').add(formData);

        // Optionally update related collections
        const schoolQuery = await db.collection('schools')
            .where('name', '==', formData.schoolName)
            .where('coordinatorId', '==', formData.coordinatorId)
            .get();

        if (!schoolQuery.empty) {
            const schoolDoc = schoolQuery.docs[0];
            await schoolDoc.ref.update({
                status: 'completed',
                workshopDate: formData.workshopDate,
                updatedAt: admin.firestore.FieldValue.serverTimestamp()
            });
        }

        res.status(200).json({ 
            message: 'Workshop Summary submitted successfully!',
            docId: docRef.id 
        });
    } catch (error) {
        console.error('Error submitting workshop summary:', error);
        res.status(500).json({ 
            error: 'Failed to submit Workshop Summary form',
            details: error.message 
        });
    }
});

// Validation function
function validateFormData(formData) {
    const errors = [];

    // Required fields validation
    const requiredFields = [
        'schoolName', 
        'schoolAddress', 
        'workshopDate', 
        'studentFeedbackFormNumber',
        'trainer1', 
        'coordinatorName', 
        'principalName', 
        'financialStatus', 
        // 'kitPaymentStatus', 
        'trainerRemunerationStatus', 
        'paymentMode', 
        'coordinatorDate', 
        'coordinatorPlace',
        'principalFeedback',
        'civicTeachersFeedback'
    ];

    requiredFields.forEach(field => {
        if (!formData[field] || (typeof formData[field] === 'string' && formData[field].trim() === '')) {
            errors.push(`${field} is required`);
        }
    });

    // Validate studentFeedbackFormNumber
    if (isNaN(formData.studentFeedbackFormNumber) || formData.studentFeedbackFormNumber < 0) {
        errors.push('Student Feedback Form Number must be a valid number (0 or greater)');
    }

    // Validate coordinatorDeclaration
    if (!formData.coordinatorDeclaration) {
        errors.push('Coordinator Declaration must be checked');
    }

    // Validate payment details for Online or Cheque
    if (formData.paymentMode === 'Online' || formData.paymentMode === 'Cheque') {
        if (!formData.upiId || formData.upiId.trim() === '') {
            errors.push('UPI ID is required for Online or Cheque payments');
        }
        if (!formData.transactionId || formData.transactionId.trim() === '') {
            errors.push('Transaction ID is required for Online or Cheque payments');
        }
    }

    // Validate feedback dropdowns
    if (formData.principalFeedback && !['Yes', 'No'].includes(formData.principalFeedback)) {
        errors.push('Principal Feedback must be either Yes or No');
    }
    if (formData.civicTeachersFeedback && !['Yes', 'No'].includes(formData.civicTeachersFeedback)) {
        errors.push('Civic Teachers Feedback must be either Yes or No');
    }

    return errors;
}



// GET endpoint for fetching workshop summaries
app.get('/workshop-summaries', isAuthenticated, async (req, res) => {
    try {
        const snapshot = await db.collection('workshopSummaries')
            .where('coordinatorId', '==', req.session.coordinator.id)
            .orderBy('createdAt', 'desc')
            .get();

        if (snapshot.empty) {
            return res.status(200).json({ message: 'No workshop summaries found', data: [] });
        }

        const summaries = snapshot.docs.map(doc => ({
            id: doc.id,
            ...doc.data(),
            createdAt: doc.data().createdAt ? doc.data().createdAt.toDate().toISOString() : null
        }));

        res.status(200).json({ data: summaries });
    } catch (error) {
        console.error('Error fetching workshop summaries:', error);
        res.status(500).json({ error: 'Failed to fetch workshop summaries' });
    }
});  
    
// Logout
app.get('/logout', (req, res) => {
    req.session.destroy(err => {
        if (err) {
            console.error('Error destroying session:', err.message, err.stack);
        }
        res.redirect('/login-coordinator');
    });
});


// Route for adding call log
// Add call log
app.post('/add-call-log', async (req, res) => {
    if (!req.body) {
        return res.status(400).json({ error: 'Request body is missing' });
    }

    const { callDate, caller, recipient, notes, schoolName, contactPerson } = req.body;

    // Check required fields
    if (!callDate || !caller || !recipient || !notes || !schoolName || !contactPerson) {
        return res.status(400).json({
            error: 'All fields (callDate, caller, recipient, notes, schoolName, contactPerson) are required'
        });
    }

    try {
        // Create new log object
        const newLog = {
            callDate,
            caller,
            recipient,
            notes,
            schoolName,
            contactPerson,
            duration: 'N/A',
            createdAt: new Date()
        };

        // Save to Firestore
        const docRef = await db.collection('callLogs').add(newLog);

        console.log('New Call Log Added:', newLog);
        res.json({ message: 'Call log added successfully', id: docRef.id });
    } catch (error) {
        console.error('Error adding call log:', error);
        res.status(500).json({ error: 'Failed to add call log' });
    }
});

// Fetch all call logs
app.get('/call-logs', async (req, res) => {
    try {
        const snapshot = await db.collection('callLogs').orderBy('callDate', 'desc').get();

        const logs = snapshot.docs.map(doc => ({
            id: doc.id,
            ...doc.data()
        }));

        res.json(logs);
    } catch (err) {
        console.error('❌ Error fetching call logs:', err);
        res.status(500).json({ error: 'Failed to fetch call logs' });
    }
});


// Express route to add login log
// Add login log
app.post('/add-login-log', requireAdmin, async (req, res) => {
    try {
        const { ipAddress, latitude, longitude, loginTime, userId } = req.body;
        if (!ipAddress || !latitude || !longitude || !loginTime) {
            return res.status(400).json({ error: 'Missing required fields' });
        }
        const parsedLatitude = typeof latitude === 'string' && latitude !== 'N/A' ? parseFloat(latitude) : latitude;
        const parsedLongitude = typeof longitude === 'string' && longitude !== 'N/A' ? parseFloat(longitude) : longitude;
        if (parsedLatitude !== 'N/A' && (typeof parsedLatitude !== 'number' || parsedLatitude < -90 || parsedLatitude > 90)) {
            return res.status(400).json({ error: 'Invalid latitude' });
        }
        if (parsedLongitude !== 'N/A' && (typeof parsedLongitude !== 'number' || parsedLongitude < -180 || parsedLongitude > 180)) {
            return res.status(400).json({ error: 'Invalid longitude' });
        }
        const ipRegex = /^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$/;
        if (!ipRegex.test(ipAddress)) {
            return res.status(400).json({ error: 'Invalid IP address' });
        }
        const parsedLoginTime = new Date(loginTime);
        if (isNaN(parsedLoginTime.getTime())) {
            return res.status(400).json({ error: 'Invalid loginTime format' });
        }
        const newLog = {
            ipAddress: ipAddress || 'N/A',
            latitude: parsedLatitude || 'N/A',
            longitude: parsedLongitude || 'N/A',
            loginTime: parsedLoginTime,
            userId: userId || req.user?.uid || 'N/A'
        };
        const docRef = await db.collection('adminLoginLogs').add(newLog);
        const logsSnapshot = await db.collection('adminLoginLogs')
            .orderBy('loginTime', 'desc')
            .get();
        const logs = logsSnapshot.docs.map(doc => ({
            id: doc.id,
            ...doc.data(),
            loginTime: doc.data().loginTime?.toDate?.()?.toISOString?.() || doc.data().loginTime
        }));
        res.status(200).json({
            message: 'Login log added successfully',
            log: { id: docRef.id, ...newLog, loginTime: newLog.loginTime.toISOString() },
            logs
        });
    } catch (error) {
        console.error('Error in /add-login-log:', error.message);
        res.status(500).json({ error: 'Failed to add login log', details: error.message });
    }
});
// Express route to fetch all login logs
app.get('/get-login-logs', async (req, res) => {
  try {
    const logsSnapshot = await db.collection('adminLoginLogs').orderBy('loginTime', 'desc').get();
    const logs = logsSnapshot.docs.map(doc => ({
      id: doc.id,
      ...doc.data()
    }));
    res.status(200).json({ logs });
  } catch (error) {
    console.error('Error fetching login logs:', error);
    res.status(500).json({ error: 'Failed to fetch login logs', details: error.message });
  }
});




// server.js or app.js
app.get('/api/impact-stats', async (req, res) => {
  try {
    // Fetch schools
    const schoolsSnapshot = await db.collection('schools').get();
    const schools = schoolsSnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    const approvedSchools = schools.filter(s => s.isApproved).length;

    // Fetch trainers
    const trainersSnapshot = await db.collection('trainers').get();
    const trainers = trainersSnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));

    // Fetch participants
    const participantsSnapshot = await db.collection('participants').get();
    const participants = participantsSnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    const studentsCompletedMCQ = participants.filter(p => p.hasCompletedMCQ).length;
    const averageScore = participants.filter(p => p.hasCompletedMCQ).length > 0
      ? (participants.reduce((sum, p) => sum + (Number(p.score) || 0), 0) / participants.filter(p => p.hasCompletedMCQ).length).toFixed(2)
      : 'N/A';

    // District-wise distribution (aggregate by city)
    const districtCounts = schools.reduce((acc, school) => {
      const city = school.city || 'Other';
      acc[city] = (acc[city] || 0) + 1;
      return acc;
    }, {});
    const districtData = Object.entries(districtCounts).map(([city, count]) => ({ city, count }));

    // Profession-wise trainers
    const professionCounts = trainers.reduce((acc, trainer) => {
      const profession = trainer.profession || 'Other';
      acc[profession] = (acc[profession] || 0) + 1;
      return acc;
    }, {});
    const professionData = Object.entries(professionCounts).map(([profession, count]) => ({ profession, count }));

    // Today's stats
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1);
    const studentsTodaySnapshot = await db.collection('participants')
      .where('hasCompletedMCQ', '==', true)
      .where('completedAt', '>=', today)
      .where('completedAt', '<', tomorrow)
      .get();
    const studentsTodayCount = studentsTodaySnapshot.size;
    const totalScoreToday = studentsTodaySnapshot.docs.reduce((sum, doc) => sum + (Number(doc.data().score) || 0), 0);
    const averageScoreToday = studentsTodayCount > 0 ? (totalScoreToday / studentsTodayCount).toFixed(2) : 0;

    res.json({
      totalSchools: schools.length,
      approvedSchools,
      totalTrainers: trainers.length,
      totalStudents: participants.length,
      studentsCompletedMCQ,
      averageScore,
      studentsTodayCount,
      averageScoreToday,
      districtData,
      professionData
    });
  } catch (err) {
    console.error('Error fetching impact stats:', err);
    res.status(500).json({ error: 'Failed to load stats' });
  }
});


// Start the server
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);

    });	
