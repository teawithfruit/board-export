'use strict';

module.exports = function() {

  var fs = require('fs');
  var excel = require('excel-export');
  var Q = require('q');
  var path = require('path');

  var JiraApi = require('jira').JiraApi;
  var jira = new JiraApi('https', config.host, config.port, config.user, config.password, '2');

  var fields = {};
  fields.cols = require('./cols');
  fields.stylesXmlFile = 'styles.xml';

  var promises = [];

  var getIssue = function(id) {
    var deferred = Q.defer();

    jira.findIssue(id, function(error, issue) {
      if(error) deferred.reject(error);
      deferred.resolve(issue);
    });
    
    return deferred.promise;
  };

  var getProject = function(id) {
    var deferred = Q.defer();

    jira.getProject(id, function(error, data) {
      if(error) deferred.reject(error);
      deferred.resolve(data);
    });
  
    return deferred.promise;
  };

  var getProjects = function() {
    var deferred = Q.defer();

    jira.listProjects(function(error, data) {
      if(error) deferred.reject(error);
      deferred.resolve(data);
    });

    return deferred.promise;
  };

  var getItems = function(name) {
    var deferred = Q.defer();
    var result = [];
    var cards = [];
    var issuePromises = [];

    jira.searchJira('project = ' + name, null, function(error, data) {
      if(error) deferred.reject(error);

      data.issues.forEach(function(tupel) {
        cards.push(tupel);
      });

      for(var i in cards) {
        issuePromises.push(getIssue(cards[i].id));
      }

      Q.allSettled(issuePromises)
      .then(function(data) {

        data.forEach(function(tupel) {
          for(var i in cards) {
            if(cards[i].id == tupel.value.id) cards[i].created = tupel.value.fields.created;
            if(cards[i].id == tupel.value.id && tupel.value.fields.status.statusCategory.key == 'done') cards[i].ready = tupel.value.fields.updated;
          }
        });

        result.push( { 'listname': name, 'cards': cards } );

        deferred.resolve(result);
      });

    });

    return deferred.promise;
  };

  var getData = function() {
    var deferred = Q.defer();

    getProjects()
    .then(function(data) {
      data.forEach(function(p) {
        promises.push(getItems(p.key));
      });

      Q.allSettled(promises)
      .then(function(data) {
        var listname = undefined;
        var blank = false;

        fields.rows.push(['', '', '', '', '', '', '', '', '']);

        data.forEach(function(tupel) {
          if(listname != '') listname = tupel.value[0].listname;

          tupel.value[0].cards.forEach(function(d) {
            var displayName = null;
            if(d.fields.assignee) displayName = d.fields.assignee.displayName;

            var ready = null;
            if(d.ready) ready = d.ready;

            fields.rows.push(
              [
                '',
                listname,
                '',
                d.fields.summary,
                '',
                d.created,
                ready,
                displayName,
                d.fields.status.statusCategory.key
              ]);

            listname = '';

            blank = true;
          });

          listname = undefined;
          if(blank == true) fields.rows.push(['', '', '', '', '', '', '', '', '']);
          blank = false;
        });

        deferred.resolve();
      });
    });

    return deferred.promise;
  };

  return {

    export: function(opts) {
      var sinceDate = new Date();
      var outfile = opts.outfile;
      var query = {
        since: null,
        before: null,
        filter: null
      };

      if(opts.since != null) {
        if(Number.isInteger(opts.since)) {
          sinceDate.setDate(sinceDate.getDate() - opts.since);
          query.since = sinceDate.getFullYear() + '-' + ("0" + (sinceDate.getMonth() + 1)).slice(-2) + '-' + ("0" + sinceDate.getDate()).slice(-2);
        } else if(Date.parse(opts.since)) {
          sinceDate = new Date(opts.since);
          query.since = sinceDate.getFullYear() + '-' + ("0" + (sinceDate.getMonth() + 1)).slice(-2) + '-' + ("0" + sinceDate.getDate()).slice(-2);
        }

        if(opts.type) {
          query.filter = opts.type;
        }
      } else {
        query.filter = opts.type;
      }

      fields.rows = [];

      getData(opts)
      .then(function() {
        var result = excel.execute(fields);
        if(!outfile) outfile = path.resolve(__dirname, 'exp_jira.xlsx');
        fs.writeFile(outfile, result, 'binary', function(err) {
          if(err) {
            return console.log(err);
          }

          console.log('all done');
          return 1;
        });
      });
      
    }

  };
}();
