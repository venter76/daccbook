// ensureAdmin.js

function ensureAdmin(req, res, next) {
    if (req.user && req.user.role === 'admin') {
      return next();
    }
  
    req.flash('error', 'Access denied');
    return res.redirect('/homedashboard');// Redirect or show an error message if the user is not an admin
   
  }
  
  module.exports = ensureAdmin;
  