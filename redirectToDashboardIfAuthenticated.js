function redirectToDashboardIfAuthenticated(req, res, next) {
    console.log("Checking if user is already authenticated");
    if (req.session && req.session.isLoggedIn) {
      return res.redirect('/homedashboard');
    }
    next();
  }
  module.exports = redirectToDashboardIfAuthenticated;