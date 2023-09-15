const express = require('express');
const bodyParser = require('body-parser');
const flash = require('connect-flash');
const mongoose = require("mongoose");
const { Schema } = mongoose;
const crypto = require('crypto');
const checkAuthenticated = require('./authenticate.js');
const redirectToDashboardIfAuthenticated = require('./redirectToDashboardIfAuthenticated');
// const pushToGoogleSheet = require('./googleSheetsModule');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const moment = require('moment');
const fs = require('fs');
const cron = require('node-cron');



// Google service account for localhost:
// Read the service account JSON file
// const rawData = fs.readFileSync('google-credentials.json');
// const serviceAccount = JSON.parse(rawData);

// Google service account for production:
// const serviceAccount = JSON.parse(process.env.GOOGLE_CREDENTIALS);





const MongoStore = require('connect-mongo');
const passport = require("passport");
const passportLocalMongoose = require("passport-local-mongoose");


const ensureAdmin = require('./ensureAdmin');

const findOrCreate = require('mongoose-findorcreate');
const schedule = require('node-schedule');
const nodemailer = require('nodemailer');
const { v4: uuidv4 } = require('uuid');

require('dotenv').config();
const exportToExceel = require('./exportToExceel');
// const createPdf = require('./createPDF');
const axios = require('axios');

 //Nodemailer setup for email verification:


const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL,
    pass: process.env.EMAIL_PASSWORD
  }
});


transporter.verify(function (error, success) {
  if(error) {
      console.log(error);
  } else {
      console.log('Server validation done and ready for messages.')
  }
}); 

const app = express();
const PORT = process.env.PORT || 3000
app.set('view engine', 'ejs');
app.use(express.static("public"));
app.use(bodyParser.urlencoded({extended: true}));
app.use(bodyParser.json()); // for parsing application/json

const bcrypt = require('bcryptjs');
const session = require('express-session');


const db_username = process.env.DB_USERNAME;
const db_password = process.env.DB_PASSWORD;
const db_cluster_url = process.env.DB_CLUSTER_URL;
const db_name = process.env.DB_NAME;

// mongodb+srv://stephenventer47:ilovebella@cluster5.9uum0ro.mongodb.net/
const connectDB = async () => {
  try {
    const conn = await mongoose.connect(`mongodb+srv://${db_username}:${db_password}@${db_cluster_url}/${db_name}?retryWrites=true&w=majority`, {
      useNewUrlParser: true,
      useUnifiedTopology: true
    });

    console.log('Connected to MongoDB Atlas:', conn.connection.host);
  } catch (error) {
    console.error('Error connecting to MongoDB Atlas:', error);
    process.exit(1);
  }
};


const userSchema = new mongoose.Schema({
  email: String,
  password: String,
  verificationToken: String, // New field for verification token
  resetPasswordToken: String, // Field for password reset token
  resetPasswordExpires: Date, // Field for token expiration time
  firstname: String,
  surname: String,
  date: {
    type: Date,
    default: Date.now
  },
  role: {
    type: String,
    default: "user" 
  },
  active: {  // Adding the 'active' field
    type: Boolean,
    default: false
  }
});

userSchema.plugin(passportLocalMongoose);
userSchema.plugin(findOrCreate);

const User = new mongoose.model("User", userSchema);

module.exports = User;






// Define the consultantSchema
const consultantSchema = new mongoose.Schema({
  firstname: {
    type: String,
    required: true
  },
  surname: {
    type: String,
    required: true
  }
});

// Create a model using the consultantSchema
const Consultant = mongoose.model('Consultant', consultantSchema);

module.exports = Consultant;



// Define the registrarSchema
const registrarSchema = new mongoose.Schema({
  firstname: {
    type: String,
    required: true
  },
  surname: {
    type: String,
    required: true
  }
});

// Create a model using the consultantSchema
const Registrar = mongoose.model('Registrar', registrarSchema);

module.exports = Registrar;











app.set('trust proxy', 1);

app.use(session({
  secret: process.env.SESSION_SECRET,
  resave: false,
  saveUninitialized: true,
  store: MongoStore.create({ 
    mongoUrl: `mongodb+srv://${db_username}:${db_password}@${db_cluster_url}/${db_name}?retryWrites=true&w=majority`,
    }),
    cookie: { 
      secure: process.env.NODE_ENV === 'production',
      sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax', // must be 'none' to enable cross-site delivery
      httpOnly: true, // prevents JavaScript from making changes
      maxAge: 14 * 24 * 60 * 60 * 1000 // 2 weeks in milliseconds
    }
  }));


  app.use(passport.initialize());
  app.use(passport.session());
  app.use(flash());
  
  passport.use(User.createStrategy());
  
  passport.serializeUser(function(user, done) {
    done(null, user.id);
  });

  passport.deserializeUser(async (id, done) => {
    try {
        const user = await User.findById(id);
        done(null, user);
    } catch (err) {
        done(err, null);
    }
});




  app.use(flash());




  const rateLimit = require('express-rate-limit');

  // Define a limiter middleware for login attempts
  const loginLimiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 15, // limit each IP to 5 requests per windowMs
    message: 'Too many login attempts, please try again in 15 minutes.'
  });
  
  // Apply the rate limiter middleware to your login route
  app.use('/login', loginLimiter);




 






async function fetchAndPushDataToSheets() {
  const staffRecords = await Staff.find({ sheet: '' }).lean().exec();
  
  // Define the headers for Excel
const headers = [
  "_id",  
  "Date","Registrar Name",  "Consultant Name", "Theatre / Clinic", 
  "Pre-op Asess", "Peri-op Plan", "Clinical Knowledge", "Data Interpret", "Interest in Teaching",
  "CVC Line", "Supervision1", 
  "Arterial Line", "Supervision2", 
  "Lumbar Epidural", "Supervision3", 
  "Thoracic Epidural", "Supervision4", 
  "Spinal Anaesthesia", "Supervision5", 
  "Nerve Blocks", "Supervision6",
  "Airway management", "Supervision7", 
  "Fibre-optic Intubation", "Supervision8", 
  "Double-lumen Tube", "Supervision9", 
  "FATE skill", "Supervision10", 
  "TOE skill", "Supervision11", 
  "Pulmonary Artery Catherter", "Supervision12", 
  "Haemodynamic management", "Supervision13", 
  "Paeds IV", "Supervision14", 
  "Paeds A-line", "Supervision15", 
  "Paeds CVP", "Supervision16", 
  "Paeds Airway", "Supervision17", 
  "Paeds caudal", "Supervision18", 
  "Paeds epidural", "Supervision19", 
  "Paeds care", "Supervision20",

  "Critical Descision", "Attention", "Communication Coll", "Communication Patient", "Presentation", "Professional", "Independance", "Logistics", "Overall Impression", 
  "Positive Comments", "Critical Comments", "Red Flag Comments"
];

// Transform the data to match the Excel structure

const transformedData = staffRecords.map(obj => {
  const score = obj.scores[0] || {};

  return [
    obj._id, score.date, obj.regName, 
    score.consName, score.theatreName, 
    score.acaScore1, score.acaScore2, score.acaScore3, score.acaScore4, score.acaScore5, 
    score.technicalScore1, getScoreWord(score.techsuperSc1),
    score.technicalScore2, getScoreWord(score.techsuperSc2),
    score.technicalScore3, getScoreWord(score.techsuperSc3),
    score.technicalScore4, getScoreWord(score.techsuperSc4),
    score.technicalScore5, getScoreWord(score.techsuperSc5),
    score.technicalScore6, getScoreWord(score.techsuperSc6),
    score.technicalScore7, getScoreWord(score.techsuperSc7),
    score.technicalScore8, getScoreWord(score.techsuperSc8),
    score.technicalScore9, getScoreWord(score.techsuperSc9),
    score.technicalScore10, getScoreWord(score.techsuperSc10),
    score.technicalScore11, getScoreWord(score.techsuperSc11),
    score.technicalScore12, getScoreWord(score.techsuperSc12),
    score.technicalScore13, getScoreWord(score.techsuperSc13),
    score.technicalPScore1, getScoreWord(score.techsuperPSc1),
    score.technicalPScore2, getScoreWord(score.techsuperPSc2),
    score.technicalPScore3, getScoreWord(score.techsuperPSc3),
    score.technicalPScore4, getScoreWord(score.techsuperPSc4),
    score.technicalPScore5, getScoreWord(score.techsuperPSc5),
    score.technicalPScore6, getScoreWord(score.techsuperPSc6),
    score.technicalPScore7, getScoreWord(score.techsuperPSc7),
    score.nonScore1, score.nonScore2, score.nonScore3, score.nonScore4, score.nonScore5, score.nonScore6, score.nonScore7, score.nonScore8, score.ratingValue, obj.positiveComments, obj.negativeComments, obj.redComments
          ];
});

transformedData.unshift(headers);

  await pushToGoogleSheet(transformedData);
  await Staff.updateMany({ sheet: '' }, { sheet: 'Y' });

  console.log('Data successfully pushed to Google Sheets!');
}
















// Authentication Route code************:


app.get('/', redirectToDashboardIfAuthenticated, (req, res) => {
  res.render('home', { 
      success: req.flash('success'),
      error: req.flash('error')
  });
});







app.get('/login', (req, res) => {
  res.render('login', { 
    success: req.flash('success'),
    error: req.flash('error') 
  });
});



app.post("/login", loginLimiter, function(req, res, next) {
  passport.authenticate("local", function(err, user, info) {

    if (err) {
      console.log(err);
      return next(err); // Pass the error to the next middleware
    }

    if (!user) {
      // Authentication failed, set flash error message
      req.flash('error', 'Incorrect username or password');
      return res.redirect("/login");
    }


    if (!user.active) {  // User exists but hasn't verified their email
      console.log("User email not verified");
      req.flash('error', 'Please verify your email to complete registration. If you cannot find the email, then register again as a new user.');
      return res.redirect("/verifytoken");
    }

    // if (user.verificationToken !== null) {
    //   console.log("No user found");
    //   req.flash('error', 'Email not verified');
    //   return res.redirect("/login");
    // }
    

    req.login(user, function(err) {
      if (err) {
        console.log(err);
        return next(err); // Pass the error to the next middleware
      }

      // At this point, the user is successfully authenticated. Mark them as logged in.
      console.log("User logged in, setting session.isLoggedIn to true");
      req.session.isLoggedIn = true;

      // If it's the user's first login (indicated by no firstname), redirect to the welcome page.
      if (!user.firstname) {
        return res.redirect("/welcome");
      }

      // If it's not the user's first login, redirect to their main page.
      return res.redirect("/homedashboard");
    });   
  })(req, res, next);
});



app.get('/verifytoken', (req, res) => {
  res.render('verifytoken', { 
    success: req.flash('success'),
    error: req.flash('error') 
  });
});








app.get("/welcome", function(req, res){
  res.render("welcome");
});




app.post('/welcome', async (req, res) => {
  // Log req.session and req.user
  // console.log('req.session:', req.session);
  // console.log('req.user:', req.user);

  const { firstName, surname } = req.body;

  console.log(req.body);


  if (!req.user) {
      return res.status(400).send("You must be logged in to access this route.");
  }

  const userId = req.user._id;

  try {
      // Update the user and fetch the updated document
      const user = await User.findByIdAndUpdate(
          userId,
          {
              firstname: firstName,
              surname: surname
          },
          { new: true }
      );

      // User information has been updated successfully
      // Redirect or render the next page here
      res.redirect('/homedashboard');

  } catch (err) {
      console.error(err);
      res.status(500).send("An error occurred while updating user information.");
  }
});


app.get('/register', function(req, res) {

  console.log('Entered register GET route');
  res.render('register', { 
    success: req.flash('success'),
    error: req.flash('error') 
  });
  });




app.post('/register', async function(req, res) {
  console.log('Entered register POST route');
  // Check if passwords match
  if (req.body.password !== req.body.passwordConfirm) {
      console.log('Passwords do not match');
      req.flash('error', 'Passwords do not match.');
      return res.redirect('/register');
  }

  const passwordPattern = /^(?=.*[A-Z])(?=.*\d)[A-Za-z\d]{7,}$/;
    if (!passwordPattern.test(req.body.password)) {
        console.log('Password does not meet requirements.');
        req.flash('error', 'Password must be minimum 7 characters with at least 1 capital letter and 1 number.');
        return res.redirect("/register");
    }

    try {
      const existingUser = await User.findOne({ username: req.body.username });
      if (existingUser) {
          if (!existingUser.active) {
              // The user exists but hasn't been verified yet, remove them and allow re-registration
              await User.deleteOne({ username: req.body.username });
          } else {
              // The user is active and verified, redirect them to login
              req.flash('error', 'User already exists. Please login.');
              return res.redirect('/login');
          }
      }

  // try {
  //     const existingUser = await User.findOne({ username: req.body.username });
  //     if (existingUser) {
  //         req.flash('error', 'User already exists. Please login.');
  //         return res.redirect('/login');
  //     }

      const user = await User.register({ username: req.body.username, active: false }, req.body.password);

      // Generate a verification token
      const verificationToken = uuidv4();
      user.verificationToken = verificationToken;

      await user.save();

      // Send verification email
      const verificationLink = `${process.env.APP_URL}/verify?token=${verificationToken}`;
      const email = {
          from: 'brayroadapps@gmail.com',
          to: user.username,
          subject: 'Email Verification',
          text: `Please copy and paste this code into text box:    ${verificationToken}`,
          // text: `Please click the following link to verify your email address: ${verificationLink}`,
      };

      try {
          await transporter.sendMail(email);
          console.log('Verification email sent');
          req.flash('success', 'Verification email has been sent - check your email inbox');
          res.redirect('/verifytoken');
      } catch (mailError) {
          console.log('Error sending email:', mailError);
          req.flash('error', 'Failed to send verification email.');
          res.redirect('/register');
      }
  } catch (err) {
      console.log(err);
      req.flash('error', 'An unexpected error occurred.');
      return res.redirect('/home');
  }
});




app.post('/verify', async function(req, res) {
  console.log('Verification route entered');

  // Get the verification token from the form data
  const verificationToken = req.body.verificationToken;

  try {
      // Find the user with the matching verification token
      const user = await User.findOne({ verificationToken: verificationToken });

// Below is origional link code for clickable link from email:


// app.get('/verify', async function(req, res) {
//   console.log('Vefification route entered');
//   const verificationToken = req.query.token;

//   try {
//       // Find the user with the matching verification token
//       const user = await User.findOne({ verificationToken: verificationToken });

      if (!user) {
          // Invalid or expired token
          console.log('Token not found or expired');
          res.send('Unauthorized login');
          return res.redirect('/'); // Use 'return' to exit the function early
      }

      // Update the user's verification status
      user.active = true;
      user.verificationToken = null; // Clear the verification token

      try {
          await user.save();
          console.log('Email verified for user2555555555555');

          // Add the success message using flash
          req.flash('success', 'Email verified for user. Please login.');
          console.log('Redirecting to login page');
          res.redirect('/login');
      } catch (saveErr) {
          console.log('Error saving user:', saveErr);
          res.redirect('/');
      }
  } catch (err) {
      console.log(err);
      res.send('Unauthorized login');
      res.redirect('/');
  }
});



app.get('/forgotpassword', function(req, res) {
  let message = req.query.message;  // Extract message from the URL parameters.
  res.render('forgotpassword', { message: message });  // Pass message to the view.
});



app.post('/forgotpassword', async function(req, res, next) {
  try {
      const buf = await crypto.randomBytes(20);
      const token = buf.toString('hex');

      const user = await User.findOne({ username: req.body.username });

      if (!user) {
          console.log('No user with this email address');
          return res.redirect('/forgotpassword?message=No%20user%20registered%20with%20this%20email%20address');
      }

      user.resetPasswordToken = token;
      user.resetPasswordExpires = Date.now() + 10800000; // 3 hours
      console.log(new Date(user.resetPasswordExpires));

      await user.save(); // Use await here instead of the callback

      const mailOptions = {
          to: user.username,
          from: 'brayroadapps@gmail.com',
          subject: 'Node.js Password Reset',
          text: 'You are receiving this because you (or someone else) have requested the reset of the password for your account.\n\n' +
          'Please click on the following link, or paste this into your browser to complete the process:\n\n' +
          'http://' + req.headers.host + '/reset/' + token + '\n\n' +
          'If you did not request this, please ignore this email and your password will remain unchanged.\n'
      };

      // Convert sendMail to Promise
      const info = await new Promise((resolve, reject) => {
          transporter.sendMail(mailOptions, (error, result) => {
              if (error) reject(error);
              else resolve(result);
          });
      });

      console.log('Email sent: ' + info.response);
      return res.redirect('/forgotpassword?message=Email%20has%20been%20sent%20with%20further%20instructions');
  
  } catch (error) {
      console.error("Error occurred:", error);
      return res.redirect('/forgotpassword?message=An%20error%20occurred');
  }
});

    

app.get('/reset/:token', async function(req, res, next) {
  try {
    const user = await User.findOne({ 
      resetPasswordToken: req.params.token, 
      resetPasswordExpires: { $gt: Date.now() } 
    });

    if (!user) {
      // handle error: no user with this token, or token expired
      console.log('Password reset token is invalid or has expired.');
      return res.redirect('/forgotpassword?message=Password%20reset%20token%20is%20invalid%20or%20has%20expired');
    }

    // if user found, render a password reset form
    res.render('reset', {
      token: req.params.token,
      error: req.flash('error')
    });
  } catch (err) {
    console.error("Error occurred:", err);
    next(err); // pass the error to your error-handling middleware, if you have one
  }
});



app.post('/reset/:token', async function(req, res, next) {
  try {
    const user = await User.findOne({ 
      resetPasswordToken: req.params.token, 
      resetPasswordExpires: { $gt: Date.now() } 
    });

    if (!user) {
      console.log('Password reset token is invalid or has expired.');
      return res.redirect('/forgotpassword?message=Password%20reset%20token%20is%20invalid%20or%20has%20expired');
    }

    const passwordPattern = /^(?=.*[A-Z])(?=.*\d)[A-Za-z\d]{7,}$/;
    if (!passwordPattern.test(req.body.password)) {
        console.log('Password does not meet requirements.');
        req.flash('error', 'Password must be minimum 7 characters with at least 1 capital letter and 1 number.');
        return res.redirect("/reset");
    }

    // Check if passwords match
    if (req.body.password !== req.body.passwordConfirm) {
      console.log('Passwords do not match');
      req.flash('error', 'Passwords do not match');
      return res.redirect("/reset");
    }

    // Assuming you have an asynchronous setPassword function. 
    await user.setPassword(req.body.password);

    user.resetPasswordToken = undefined;
    user.resetPasswordExpires = undefined;

    await user.save();

    // Wrap req.logIn in a promise
    await new Promise((resolve, reject) => {
      req.logIn(user, (err) => {
        if (err) reject(err);
        else resolve();
      });
    });

    res.redirect('/');

  } catch (err) {
    console.error("Error occurred:", err);
    next(err); 
  }
});












// Othe routes code***********:



app.get('/homedashboard', checkAuthenticated, async function (req, res) {
  console.log("Entered /homedashboard route");
  try {
      const user = await User.findById(req.user._id);
      
      if (!user) {
          // If the user doesn't exist, you can handle it accordingly
          // Here, I'm sending a 404 Not Found status. You can modify as needed.
          return res.status(404).send('User not found');
      }

      res.render('homedashboard', {error: req.flash('error')});

  } catch (err) {
      console.log(err);
      // Send a generic error message or a specific one based on the nature of the error
      res.status(500).send('An error occurred while fetching user information.');
  }
});







app.get('/assess', async (req, res) => {
   // Fetch all registrars from the database
   let registrars = [];
   try {
    registrars = await Registrar.find().sort({ surname: 1 });  // Sorting by surname in ascending order
   } catch (err) {
     console.error("Failed to retrieve registrars:", err);
   }
 
   // Pass the flash message and the registrars to the ejs template
   res.render('assess.ejs', { success: req.flash('success'), registrars: registrars });
 });

 










app.post('/assess', async (req, res) => {

// Access the logged-in user's first name and surname
const loggedInUserFirstName = req.user.firstname;
const loggedInUserSurname = req.user.surname;

  let formData = req.body;

// Log the ratingValue
console.log("ratingValue:", formData.rating);



  // Handle form data

  for(let key in formData) {

  //    // Log the techsuperSc and techsuperPSc values
  //    if(key.startsWith("techsuperSc") || key.startsWith("techsuperPSc")) {
  //     console.log(`${key}: ${formData[key]}`);
  // }
    // check if the value of the key in the form data is 0 or null
    // if yes then change it to null (not the string "null")
    if(formData[key] === "0" || formData[key] === null || formData[key] === "null") {
        formData[key] = null;
    } else if (!isNaN(formData[key])) {
        // convert numerical strings to numbers
        formData[key] = Number(formData[key]);
    }
  }




  // Separate staff data and score data
  const staffData = {
    regName: formData.regName,
    // id: formData.id, // include this if you have an 'id' field
    positiveComments: formData.positiveComments,
    negativeComments: formData.negativeComments,
    redComments: formData.redComments,
  };

  const scoreData = { ...formData, ratingValue: formData.rating };
  delete scoreData.regName;
  delete scoreData.positiveComments;
  delete scoreData.negativeComments;
  delete scoreData.redComments;
  // delete scoreData.id; // include this if you have an 'id' field

// Calculate the total and the count of acaScores
let acaScoreTotal = 0;
let acaScoreCount = 0;
for (let i = 1; i <= 5; i++) {
  if (scoreData[`acaScore${i}`] !== null) {
    acaScoreTotal += scoreData[`acaScore${i}`];
    acaScoreCount++;
  }
}

// Calculate the total and the count of (adult) technicalScores
let technicalScoreTotal = 0;
let technicalScoreCount = 0;
// *****************************************************************
for (let i = 1; i <= 13; i++) {  
  if (scoreData[`technicalScore${i}`] !== null) {
    technicalScoreTotal += scoreData[`technicalScore${i}`];
    technicalScoreCount++;
  }
}

// Calculate the total and the count of (paediatric) technicalPScores
let technicalPScoreTotal = 0;
let technicalPScoreCount = 0;
// *****************************************************************
for (let i = 1; i <= 7; i++) {  
  if (scoreData[`technicalPScore${i}`] !== null) {
    technicalPScoreTotal += scoreData[`technicalPScore${i}`];
    technicalPScoreCount++;
  }
}

// Calculate the total and the count of nonScores
let nonScoreTotal = 0;
let nonScoreCount = 0;
for (let i = 1; i <= 8; i++) {
  if (scoreData[`nonScore${i}`] !== null) {
    nonScoreTotal += scoreData[`nonScore${i}`];
    nonScoreCount++;
  }
}

// Calculate the average acaScore excluding null entries and assign to acaScoretotal
if (acaScoreCount > 0) {
  scoreData.acaScoretotal = acaScoreTotal / acaScoreCount;
} else {
  scoreData.acaScoretotal = null;
}

// Calculate the average technicalScore excluding null entries and assign to technicalScoretotal
if (technicalScoreCount > 0) {
  scoreData.technicalScoretotal = technicalScoreTotal / technicalScoreCount;
} else {
  scoreData.technicalScoretotal = null;
}

// Calculate the average technicalScore excluding null entries and assign to technicalPScoretotal
if (technicalPScoreCount > 0) {
  scoreData.technicalPScoretotal = technicalPScoreTotal / technicalPScoreCount;
} else {
  scoreData.technicalPScoretotal = null;
}


// Calculate the average nonScore excluding null entries and assign to nonScoretotal
if (nonScoreCount > 0) {
  scoreData.nonScoretotal = nonScoreTotal / nonScoreCount;
} else {
  scoreData.nonScoretotal = null;
}

// Create a new staff document
const staff = new Staff({
  ...staffData,
  scores: [
    {
      ...scoreData,
      consName: `${loggedInUserFirstName} ${loggedInUserSurname}` // Add consName to each scoreData object
    }
  ]
});

// Save the staff document

try {
  const savedStaff = await staff.save();
  console.log(savedStaff);

  // Set the flash message
  req.flash('success', 'Data submitted');

  // Retrieve the flash message only once
  const successMessage = req.flash('success');
  
  // Log the success message
  console.log("Success message:", successMessage);

  let registrars = [];
try {
  registrars = await Registrar.find();
} catch (err) {
  console.error("Failed to retrieve registrars:", err);
}

res.render('assess.ejs', { success: successMessage, registrars: registrars });


} catch (err) {
  console.error(err);
  req.flash('error', 'Error saving data');
  res.redirect('/assess'); // redirect back to the assess page or another error page
}



});









app.post('/dateList2', async (req, res) => {
  try {
      const selectedDate = new Date(req.body.dateInput);
      console.log(selectedDate);

      const results = await Staff.find({ "scores.date": selectedDate }).exec();

      const data = results.reduce((acc, doc) => {
          doc.scores.forEach(score => {
              if (score.date && score.date.toISOString() === selectedDate.toISOString()) {
                  acc.push({
                      regName: doc.regName,
                      theatreName: score.theatreName
                  });
              }
          });
          return acc;
      }, []);

      console.log(data);

      // Render the finalList.ejs page and pass the data
      res.render('finalList', { date: selectedDate, data: data });

  } catch (err) {
      console.error("There was an error fetching the data:", err);
      res.status(500).send('Server Error');
  }
});





app.get('/admin', checkAuthenticated, ensureAdmin, (req, res) => {
// Pass the flash message to the ejs template
res.render('admin.ejs', { success: req.flash('success') });
});








app.get('/howto', (req, res) => {
  res.render('howto');
  
});








app.get('/report', checkAuthenticated, ensureAdmin, async (req, res) => {
   // Fetch all registrars from the database
   let registrars = [];
   try {
    registrars = await Registrar.find().sort({ surname: 1 });  // Sorting by surname in ascending order
   } catch (err) {
     console.error("Failed to retrieve registrars:", err);
   }
 
   // Pass the flash message and the registrars to the ejs template
   res.render('report.ejs', { registrars: registrars, error: req.flash('error') });
 });

 







// PLACE NAME ADJUSTERS *************************************************

app.post('/reportdig', async (req, res) => {
   
  let regName = req.body.regName;

  // Check if date values are provided before creating Date objects
  let dateFrom = req.body.dateFrom ? new Date(req.body.dateFrom) : null;
  let dateTo = req.body.dateTo ? new Date(req.body.dateTo) : null;

  console.log("regName:", regName);
  console.log("dateFrom:", dateFrom);
  console.log("dateTo:", dateTo);
  
  try {
    // Create a base query and then conditionally add date range
    let query = { regName: regName };
    
    if (dateFrom && dateTo) {
      query["scores.date"] = { $gte: dateFrom, $lte: dateTo };
    }

    const staffData = await Staff.find(query);

    if (!staffData || staffData.length === 0) {
      console.log(`No staff found with name: ${regName}`);
      
      // Set the flash message
      req.flash('error', `No data for selected dates for: ${regName}`);
      
      
      return res.redirect('/report');
    }



    // if (!staffData || staffData.length === 0) {
    //   console.log(`No staff found with regName: ${regName}`);
    //   return res.status(404).send(`No staff found with regName: ${regName}`);
    // }

    let scoresData = []; // This will contain all the score data for each staff member
    let positiveCommentsArray = [];
    let negativeCommentsArray = [];
    let redCommentsArray = [];


    // console.log("Entering staffData loop...");

    staffData.forEach(doc => {

      // console.log("Processing a staff document...");
      // console.log("Current doc:", JSON.stringify(doc, null, 2));

      // Push the comments from this document to the respective arrays
      if (doc.positiveComments && doc.positiveComments !== '0') positiveCommentsArray.push(doc.positiveComments);
      if (doc.negativeComments && doc.negativeComments !== '0') negativeCommentsArray.push(doc.negativeComments);
      if (doc.redComments && doc.redComments !== '0') redCommentsArray.push(doc.redComments);

      doc.scores.forEach(score => {

    


          // Construct an object with only the needed data
          let scoreData = {
              date: score.date,
              acaScoretotal: score.acaScoretotal,
              technicalScoretotal: score.technicalScoretotal,
              technicalPScoretotal: score.technicalPScoretotal,
              nonScoretotal: score.nonScoretotal,
              ratingValue: score.ratingValue,
          };
          scoresData.push(scoreData);
      });
    });
  
    // Sort scoresData based on the date field
    scoresData.sort((a, b) => new Date(a.date) - new Date(b.date));
  
    // console.log(scoresData);

    let acaScoreSum = scoresData.reduce((accum, score) => accum + score.acaScoretotal, 0);
    let technicalScoreSum = scoresData.reduce((accum, score) => accum + score.technicalScoretotal, 0);
    let technicalPScoreSum = scoresData.reduce((accum, score) => accum + score.technicalPScoretotal, 0);
    let nonScoreSum = scoresData.reduce((accum, score) => accum + score.nonScoretotal, 0);

    let scoresLength = scoresData.length || 1;  // Avoid division by zero
    let acaScoreAverage = acaScoreSum / scoresLength;
    let technicalScoreAverage = technicalScoreSum / scoresLength;
    let technicalPScoreAverage = technicalPScoreSum / scoresLength;
    let nonScoreAverage = nonScoreSum / scoresLength;

    // console.log('Average Academic Score:', acaScoreAverage);
    // console.log('Average Technical Score:', technicalScoreAverage);
    // console.log('Average Paeds Technical Score:', technicalPScoreAverage);
    // console.log('Average Non-Technical Score:', nonScoreAverage);

    const analyzedScores = analyzeScores(staffData);
    // console.log(analyzedScores);
    

    
   

const scoreNameMapping = {
  acaScore1: "Pre-op Assessment",
  acaScore2: "Peri-op Management",
  acaScore3: "Academic / Clinical Knowledge",
  acaScore4: "Data Interpretation",
  acaScore5: "Interest in Teaching",
  technicalScore1: "CVP",
    technicalScore2: "Arterial Line",
    technicalScore3: "Lumbar Epidural",
    technicalScore4: "Thoracic Epidural",
    technicalScore5: "Spinal Anaesthesia",
    technicalScore6: "Nerve Block",
    technicalScore7: "Airway management",
    technicalScore8: "Fibreoptic intubation",
    technicalScore9: "Double-lumen Tube",
    technicalScore10: "FATE skills",
    technicalScore11: "TOE skills",
    technicalScore12: "Pulmonary Artery Catheter",
    technicalScore13: "Haemodynamic Management",

    technicalPScore1: "IV cannula",
    technicalPScore2: "Arterial Line",
    technicalPScore3: "CVP",
    technicalPScore4: "Airway management",
    technicalPScore5: "Caudal anaesthesia",
    technicalPScore6: "Epidural anaesthesia",
    technicalPScore7: "General Care",

    nonScore1: "Critical Descision Making",
    nonScore2: "Attention to Detail",
    nonScore3: "Communication with Colleagues",
    nonScore4: "Communication with Patients",
    nonScore5: "Presentation Skills",
    nonScore6: "Professionalism",
    nonScore7: "Independance",
    nonScore8: "Logistics/Organisational Skills",

    ratingValue: "Overall Impression",
};
 

const humanReadableScores = {};

for (let key in analyzedScores) {
    const newName = scoreNameMapping[key] || key;  
    humanReadableScores[newName] = analyzedScores[key];
}

// Console log the mapped scores
// console.log('Mapped Scores:', humanReadableScores);


// Now, pass humanReadableScores to your view (ejs)

  
    // Render the page and pass the data
    res.render('reportdig', {
      regName: regName,
      scoresData: scoresData,
      acaScoreAverage: acaScoreAverage,
      technicalScoreAverage: technicalScoreAverage,
      technicalPScoreAverage: technicalPScoreAverage,
      nonScoreAverage: nonScoreAverage,
      positiveComments: positiveCommentsArray,
      negativeComments: negativeCommentsArray,
      redComments: redCommentsArray,
      analyzedData: humanReadableScores
    });

  } catch(err) {
    console.error(err);
    res.status(500).send('Server Error');
  }
});









// *****************************************************************


// Route to push to Google SHeets (this is run automatically with node-cron)
app.get('/download1', async (req, res) => {
  try {
    await fetchAndPushDataToSheets();
    res.send('Data successfully pushed to Google Sheets!');
  } catch (error) {
    console.error('Error pushing data to Google Sheets:', error);
    res.status(500).send('Error pushing data to Google Sheets');
  }
});


















// This route to download Excel data for a particular reg +- time period

app.get('/download2', async (req, res) => {
  try {
    const regName = req.query.regName;

    // Check if regName exists
    if (!regName) {
      return res.status(400).send("Parameter 'regName' is missing");
    }

    const staffRecords = await Staff.find({ regName: regName }).lean().exec();

    // Define the headers for Excel
    const headers = [
      "_id",  
      "Date","Registrar Name",  "Consultant Name", "Theatre / Clinic", 
      "Pre-op Asess", "Peri-op Plan", "Clinical Knowledge", "Data Interpret", "Interest in Teaching",
      "CVC Line", "Supervision", 
      "Arterial Line", "Supervision", 
      "Lumbar Epidural", "Supervision", 
      "Thoracic Epidural", "Supervision", 
      "Spinal Anaesthesia", "Supervision", 
      "Nerve Blocks", "Supervision",
      "Airway management", "Supervision", 
      "Fibre-optic Intubation", "Supervision", 
      "Double-lumen Tube", "Supervision", 
      "FATE skill", "Supervision", 
      "TOE skill", "Supervision", 
      "Pulmonary Artery Catherter", "Supervision", 
      "Haemodynamic management", "Supervision", 
      "Paeds IV", "Supervision", 
      "Paeds A-line", "Supervision", 
      "Paeds CVP", "Supervision", 
      "Paeds Airway", "Supervision", 
      "Paeds caudal", "Supervision", 
      "Paeds epidural", "Supervision", 
      "Paeds care", "Supervision",

      "Critical Descision", "Attention", "Communication Coll", "Communication Patient", "Presentation", "Professional", "Independance", "Logistics", "Overall Impression", 
      "Positive Comments", "Critical Comments", "Red Flag Comments"
    ];
    
    // Transform the data to match the Excel structure

    const transformedData = staffRecords.map(obj => {
      const score = obj.scores[0] || {};
  
      

      return [
        obj._id, score.date, obj.regName, 
        score.consName, score.theatreName, 
        score.acaScore1, score.acaScore2, score.acaScore3, score.acaScore4, score.acaScore5, 
        score.technicalScore1, getScoreWord(score.techsuperSc1),
        score.technicalScore2, getScoreWord(score.techsuperSc2),
        score.technicalScore3, getScoreWord(score.techsuperSc3),
        score.technicalScore4, getScoreWord(score.techsuperSc4),
        score.technicalScore5, getScoreWord(score.techsuperSc5),
        score.technicalScore6, getScoreWord(score.techsuperSc6),
        score.technicalScore7, getScoreWord(score.techsuperSc7),
        score.technicalScore8, getScoreWord(score.techsuperSc8),
        score.technicalScore9, getScoreWord(score.techsuperSc9),
        score.technicalScore10, getScoreWord(score.techsuperSc10),
        score.technicalScore11, getScoreWord(score.techsuperSc11),
        score.technicalScore12, getScoreWord(score.techsuperSc12),
        score.technicalScore13, getScoreWord(score.techsuperSc13),
        score.technicalPScore1, getScoreWord(score.techsuperPSc1),
        score.technicalPScore2, getScoreWord(score.techsuperPSc2),
        score.technicalPScore3, getScoreWord(score.techsuperPSc3),
        score.technicalPScore4, getScoreWord(score.techsuperPSc4),
        score.technicalPScore5, getScoreWord(score.techsuperPSc5),
        score.technicalPScore6, getScoreWord(score.techsuperPSc6),
        score.technicalPScore7, getScoreWord(score.techsuperPSc7),
        score.nonScore1, score.nonScore2, score.nonScore3, score.nonScore4, score.nonScore5, score.nonScore6, score.nonScore7, score.nonScore8, score.ratingValue, obj.positiveComments, obj.negativeComments, obj.redComments
              ];
    });

    transformedData.unshift(headers);


    // wsData.unshift(headers);  // Add headers at the start of the data array
 
    // Pass the transformed data to the export function

    const buffer = await exportToExceel(transformedData);

    // const buffer = await exportToExceel(wsData);

    // Generate a timestamp
    const now = new Date();
    const sanitizedRegName = regName.replace(/[^a-z0-9]/gi, '_').toLowerCase(); // To ensure a safe filename
    const timestamp = `${sanitizedRegName}_${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()}_${now.getHours()}-${now.getMinutes()}-${now.getSeconds()}`;

    

    // Set up response headers
    res.setHeader('Content-Disposition', `attachment; filename="output_${timestamp}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Send buffer to client to trigger file download
    res.send(buffer);
  } catch (error) {
    console.error('Error exporting data to Excel:', error);
    res.status(500).send('Error exporting data to Excel');
  }
});







// Handle the POST request for '/consmanage'
app.post('/consmanage', async (req, res) => {
  const { firstname, surname, action } = req.body;

  if (action === 'add') {
    try {
      await Consultant.create({ firstname, surname });
      req.flash('success', 'New consultant added successfully');
    } catch (error) {
      console.error(error);
      req.flash('success', 'Error adding consultant');
    }
  } else if (action === 'delete') {
    try {
      await Consultant.findOneAndDelete({ firstname, surname });
      req.flash('success', 'Consultant deleted successfully');
    } catch (error) {
      console.error(error);
      req.flash('success', 'Error deleting consultant');
    }
  }

  res.redirect('/admin');  // Assuming '/admin' is the route to render admin.ejs
});








// Handle the POST request for '/regmanage'
app.post('/regmanage', async (req, res) => {
  const { firstname, surname, action } = req.body;

  if (action === 'add') {
    try {
      await Registrar.create({ firstname, surname });
      req.flash('success', 'New registrar added successfully');
    } catch (error) {
      console.error(error);
      req.flash('success', 'Error adding registrar');
    }
  } else if (action === 'delete') {
    try {
      await Registrar.findOneAndDelete({ firstname, surname });
      req.flash('success', 'Registrar deleted successfully');
    } catch (error) {
      console.error(error);
      req.flash('success', 'Error deleting registrar');
    }
  }

  res.redirect('/admin');  // Assuming '/admin' is the route to render admin.ejs
});







app.get('/conslist', async (req, res) => {

// Fetch all consultants from the database
let consultants = [];
try {
 consultants = await Consultant.find().sort({ surname: 1 });  // Sorting by surname in ascending order
} catch (err) {
  console.error("Failed to retrieve consultants:", err);
}

  // try {
  //     // Retrieve all consultants from the database
  //     const consultants = await Consultant.find({});

      // Render the conslist.ejs template and pass the consultants data
      res.render('conslist.ejs', { consultants: consultants });
  // } catch (err) {
  //     console.error("Error fetching consultants:", err);
  //     res.status(500).send("Internal Server Error");
  // }
});





app.get('/reglist', async (req, res) => {

 // Fetch all registrars from the database
 let registrars = [];
 try {
  registrars = await Registrar.find().sort({ surname: 1 });  // Sorting by surname in ascending order
 } catch (err) {
   console.error("Failed to retrieve registrars:", err);
 }

      // // Retrieve all registrars from the database
      // const registrars = await Registrar.find({});

      // Render the registrar.ejs template and pass the consultants data
      res.render('reglist.ejs', { registrars: registrars });
  // } catch (err) {
  //     console.error("Error fetching registrars:", err);
  //     res.status(500).send("Internal Server Error");
  // }
});





// app.get('/pdf', async (req, res) => {
//   const { url } = req.query;

//   if (!url) {
//     return res.status(400).send('URL is required');
//   }

//   try {
//     // Generate PDF using printPDF function
//     const pdf = await createPdf(url);

//     // Set the response headers to specify the PDF content type and disposition.
//     res.set({ 'Content-Type': 'application/pdf', 'Content-Length': pdf.length });

//     // Send the PDF as the response
//     res.send(pdf);
//   } catch (error) {
//     console.error(error);
//     res.status(500).send('An error occurred while generating the PDF');
//   }
// });

 

cron.schedule('0 * * * *', async () => {
  console.log('Running the task every 1 hours');
  try {
      await fetchAndPushDataToSheets();
  } catch (error) {
      console.error('Error in cron job:', error);
  }
});











connectDB().then(() => {
    app.listen(PORT, () => {
        console.log("listening for requests");
    })
  })



  // async function createOrUpdateInfo() {

//   const enter = process.env.ENTER;
//   const password = process.env.PASSWORD;

  


//   // Hash the password
//   const salt = await bcrypt.genSalt(10);
//   const hashedPassword = await bcrypt.hash(password, salt);

//   // Check if Info collection already has an entry
//   const info = await Info.findOne({});

//   if (info) {
//     // Info collection has an entry, update it
//     info.enter = enter;
//     info.password = hashedPassword;
//     await info.save();
//   } else {
//     // Info collection is empty, create a new entry
//     const newInfo = new Info({ enter, password: hashedPassword });
//     await newInfo.save();
//   }

//   console.log('Username and password have been saved to the database.');
// }

// // Run the function
// createOrUpdateInfo().catch(console.error);
