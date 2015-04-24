'use strict';

var trello = require("./lib/trello");
var jira = require("./lib/jira");
var parser = require("nomnom");
var path = require("path");

parser.command('trello')
  .option('outfile', {
    abbr: 'o',
    default: path.resolve(__dirname, 'exp_trello.xlsx'),
    help: "file to write results to"
  })
  .option('since', {
    abbr: 's',
    default: null,
    help: "entries since date"
  })
  .option('type', {
    abbr: 't',
    default: 'all',
    help: "get entries by type"
  })
  .callback(function(opts) {
    trello.export(opts);
  })
  .help("exports trello to excel");

parser.command('jira')
  .option('outfile', {
    abbr: 'o',
    default: path.resolve(__dirname, 'exp_jira.xlsx'),
    help: "file to write results to"
  })
  .option('since', {
    abbr: 's',
    default: null,
    help: "entries since date"
  })
  .option('type', {
    abbr: 't',
    default: 'all',
    help: "get entries by type"
  })
  .callback(function(opts) {
    jira.export(opts);
  })
  .help("exports trello to excel");

parser.parse();

// node index.js trello --type=all --since=30 --outfile=exp_trello_all.xlsx && open exp_trello_all.xlsx
// node index.js trello --type=all --outfile=exp_trello_all_all.xlsx && open exp_trello_all_all.xlsx
// node index.js trello --type=closed --since=30 --outfile=exp_trello_closed.xlsx && open exp_trello_closed.xlsx
// node index.js trello --type=all --since=2015-04-20 && open exp_trello.xlsx
// node index.js trello --type=closed --since=2015-04-20 && open exp_trello.xlsx

// node index.js jira --type=all --since=30 --outfile=exp_jira_all.xlsx && open exp_jira_all.xlsx
// node index.js jira --type=all --outfile=exp_jira_all_all.xlsx && open exp_jira_all_all.xlsx
// node index.js jira --type=closed --since=30 --outfile=exp_jira_closed.xlsx && open exp_jira_closed.xlsx
// node index.js jira --type=all --since=2015-04-20 && open exp_jira.xlsx
// node index.js jira --type=closed --since=2015-04-20 && open exp_jira.xlsx