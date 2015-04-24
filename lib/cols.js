module.exports = [
  {
    caption: 'Projektnr.',
    captionStyleIndex: 1,
    type: 'string',
    width: 9
  },
  {
    caption: 'Projektbezeichnung',
    captionStyleIndex: 1,
    type: 'string',
    width: 16
  },
  {
    caption: 'Projektbereich',
    captionStyleIndex: 1,
    type: 'string',
    width: 12
  },
  {
    caption: 'Status',
    captionStyleIndex: 1,
    type: 'string',
    width: 20
  },
  {
    caption: 'Next Steps',
    captionStyleIndex: 1,
    type: 'string',
    beforeCellWrite:function(row, cellData) {
         return cellData; //cellData.toUpperCase();
    },
    width: 9
  },
  {
    caption: 'Termin (initial)',
    captionStyleIndex: 1,
    type: 'string',
    beforeCellWrite:function(row, cellData) {
      if(cellData) {
        if(cellData.search('T') == -1) {
          var dateCreated = new Date(1000 * parseInt(cellData.substring(0,8), 16));
          return ("0" + dateCreated.getDate()).slice(-2) + '.' + ("0" + (dateCreated.getMonth() + 1)).slice(-2) + '.' + dateCreated.getFullYear();
        } else {
          var parts = cellData.split('T');
          var day = parts[0].split('-');
          var time = parts[1].split(':');

          var d = new Date(day[0], day[1], day[2], time[0], time[1], time[1].split('.')[0]); 
          
          return ("0" + d.getDate()).slice(-2) + '.' + ("0" + (d.getMonth() + 1)).slice(-2) + '.' + d.getFullYear();
        }

      }
      
      return cellData;
    },
    width: 11
  },
  {
    caption: 'Termin (aktuell)',
    captionStyleIndex: 1,
    type: 'string',
    beforeCellWrite:function(row, cellData) {
      if(cellData) {
        var parts = cellData.split('T');
        var day = parts[0].split('-');
        var time = parts[1].split(':');

        var d = new Date(day[0], day[1], day[2], time[0], time[1], time[1].split('.')[0]); 
        
        return ("0" + d.getDate()).slice(-2) + '.' + ("0" + (d.getMonth() + 1)).slice(-2) + '.' + d.getFullYear();
      }
      
      return cellData;
    },
    width: 12
  },
  {
    caption: 'Verantwortung',
    captionStyleIndex: 1,
    type: 'string',
    beforeCellWrite:function(row, cellData) {
      if(cellData) {
        var members = '';

        if(Array.isArray(cellData)) {
          cellData.forEach(function(m) {
            members = members + m.fullName + '\r';
          });

          return members;
        }
      }

      return cellData;
    },
    width: 12
  },
  {
    caption: '!',
    captionStyleIndex: 1,
    type: 'string',
    beforeCellWrite:function(row, cellData) {
      if(cellData == true || cellData == 'done') {
        return 'Abgeschlossen';
      } else {
        return '';
      }
    },
    width: 12
  }
]