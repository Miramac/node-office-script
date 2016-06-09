var application = require('./edge/application')
var app

module.exports = {
  application: application,
  open: function (path, cb) {
    if (!app) {
      app = application(null, true)
    }
    return app.open(path, cb)
  },
  quit: function (param, cb) {
    if (app) {
      app.quit(param, cb)
      app = null
    }
  },
  fetch: function (name, cb) {
    if (!app) {
      app = application(null, true)
    }
    return app.fetch(name, cb)
  }
}
