'use strict';

module.exports = function() {

  var fs = require('fs');
  var excel = require('excel-export');
  var Q = require('q');
  var path = require('path');
  var Trello = require("node-trello");
  var t = new Trello("KEY", "SECRET");

  var fields = {};
  fields.cols = require('./cols');
  fields.stylesXmlFile = 'styles.xml';

  var promises = [];

  var getLists = function(outfile, query) {
    t.get("/1/boards/4FS6rVLF/lists", { fields: 'name,pos,closed', filter: 'all' }, function(err, listdata) {
      fields.rows = [];

      for(var i in listdata) {
        promises.push(getItems(listdata[i], query));
      }

      Q.allSettled(promises)
      .then(function(data) {
        var listname = undefined;
        var blank = false;

        fields.rows.push(['', '', '', '', '', '', '', '', '']);

        data.forEach(function(tupel) {
          if(listname != '') listname = tupel.value[0].listname;

          tupel.value[0].cards.forEach(function(d) {
            
            fields.rows.push(
              [
                '',
                listname,
                '',
                d.name,
                '',
                d.id,
                d.due,
                d.members,
                d.closed,
              ]);

            listname = '';

            blank = true;
          });

          listname = undefined;
          if(blank == true) fields.rows.push(['', '', '', '', '', '', '', '', '']);
          blank = false;
        });

        var result = excel.execute(fields);
        if(!outfile) outfile = path.resolve(__dirname, 'exp_trello.xlsx');
        fs.writeFile(outfile, result, 'binary', function(err) {
          if(err) {
            return console.log(err);
          }

          console.log('all done');
          return 1;
        });

      });
    });

    Q.any(promises)
    .then(function(first) {
      return first;
    }, function(error) {
      return error;
    });
  };

  var getItems = function(data, query) {
    var deferred = Q.defer();
    var result = [];
    var cards = [];

    t.get("/1/lists/"+ data.id +"/cards", { members: true, member_fields: 'username,fullName', fields: 'idList,name,desc,idChecklists,subscribed,labels,due,closed',  filter: query.filter, since: query.since, before: query.before }, function(err, carddata) {
      for(var i in carddata) {
        cards.push( carddata[i] );
      }

      result.push( { 'listname': data.name, 'cards': cards } );

      deferred.resolve(result);
    });

    return deferred.promise;
  };

  return {

    export: function(opts) {
      var sinceDate = new Date();
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

      return getLists(opts.outfile, query);
    }

  };
}();
