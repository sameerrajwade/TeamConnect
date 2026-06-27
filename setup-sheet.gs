/**
 * Run this function ONCE to create the Three60 Google Sheet with all tabs and sample data.
 * After running, copy the Sheet ID from the logs and paste it into Code.gs.
 */
function createThree60Sheet() {
  var ss = SpreadsheetApp.create('Three60');
  Logger.log('Sheet created! ID: ' + ss.getId());
  Logger.log('URL: ' + ss.getUrl());

  // ─── EVENTS ───
  var events = ss.getSheetByName('Sheet1');
  events.setName('Events');
  events.getRange(1, 1, 7, 11).setValues([
    ['Title', 'PlannedDate', 'Time', 'HostTeam', 'Activity', 'ActivityDesc', 'MeetLink', 'Summary', 'RecordingURL', 'SlidesURL', 'NotesURL'],
    ['July Connect', new Date('2025-07-17'), '2:00 PM ET', 'DCX', 'Speed Networking', 'Random breakout pairs across all three teams', 'https://meet.google.com/abc-defg-hij', '', '', '', ''],
    ['August Blend', new Date('2025-08-21'), '2:00 PM ET', 'CAMS', 'Problem Jam', 'Bring a real challenge — get fresh perspectives', 'https://meet.google.com/abc-defg-hij', '', '', '', ''],
    ['September Spotlight', new Date('2025-09-18'), '2:00 PM ET', 'CDPS', 'Trivia Challenge', 'Cross-team trivia with prizes', 'https://meet.google.com/abc-defg-hij', '', '', '', ''],
    ['June Pilot', new Date('2025-06-19'), '2:00 PM ET', 'DCX', 'Speed Intros', 'Quick round-robin intros', '', 'Our first ever Three60 session', 'https://drive.google.com/recording1', 'https://docs.google.com/slides1', 'https://docs.google.com/notes1'],
    ['May Planning Kickoff', new Date('2025-05-15'), '2:00 PM ET', 'CAMS', 'Goal Alignment', 'Cross-team Q3 goal alignment', '', 'Aligned on Q3 priorities across all 3 teams', '', 'https://docs.google.com/slides2', 'https://docs.google.com/notes2'],
    ['April Tool Spotlight', new Date('2025-04-17'), '2:00 PM ET', 'CDPS', 'Tool Demo', 'CAMS campaign automation walkthrough', '', 'CAMS automation now used by 2 other teams', 'https://drive.google.com/recording3', 'https://docs.google.com/slides3', '']
  ]);
  events.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  events.setFrozenRows(1);
  events.autoResizeColumns(1, 11);

  // ─── SHOW AND TELL ───
  var st = ss.insertSheet('ShowAndTell');
  st.getRange(1, 1, 6, 7).setValues([
    ['Timestamp', 'PresenterName', 'Team', 'Topic', 'Description', 'Status', 'EventTitle'],
    [new Date('2025-06-10T10:00:00'), 'Priya Sharma', 'CDPS', 'How we cut pipeline runtime by 60%', 'Deep dive into our data pipeline optimization', 'Approved', 'July Connect'],
    [new Date('2025-06-15T14:30:00'), 'Marcus Thompson', 'CAMS', 'Campaign tracker in Google Sheets', 'Built a self-serve campaign tracker', 'Pending', 'August Blend'],
    [new Date('2025-05-01T09:00:00'), 'Priya Sharma', 'CDPS', 'Data pipeline rebuild', 'Walkthrough of the new pipeline architecture', 'Approved', 'June Pilot'],
    [new Date('2025-04-02T11:00:00'), 'Marcus Thompson', 'CAMS', 'Campaign automation', 'How CAMS automated their weekly campaign reporting', 'Approved', 'May Planning Kickoff'],
    [new Date('2025-03-10T15:00:00'), 'Jordan Lee', 'DCX', 'Our toolbox', 'Overview of the design tools and workflows DCX uses', 'Approved', 'April Tool Spotlight']
  ]);
  st.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  st.setFrozenRows(1);
  st.autoResizeColumns(1, 7);

  // ─── WINS ───
  var wins = ss.insertSheet('Wins');
  wins.getRange(1, 1, 9, 4).setValues([
    ['Timestamp', 'Name', 'Team', 'Description'],
    [new Date('2025-06-25T09:00:00'), 'Jordan Lee', 'DCX', 'Automated the weekly reporting workflow — saving 5 hours per week across the team'],
    [new Date('2025-06-22T14:00:00'), 'Marcus Thompson', 'CAMS', 'Led the cross-team demo that secured the Q2 partnership with the marketing analytics vendor'],
    [new Date('2025-06-18T11:00:00'), 'Priya Sharma', 'CDPS', 'Shipped the new data pipeline 2 weeks ahead of schedule — already processing 3x the previous volume'],
    [new Date('2025-06-12T10:00:00'), 'Aisha Khan', 'DCX', 'Reduced customer onboarding time by 40% through a redesigned welcome flow'],
    [new Date('2025-06-08T16:00:00'), 'Leila Hassan', 'CAMS', 'Launched the new campaign dashboard that all 3 teams are now using for real-time metrics'],
    [new Date('2025-05-30T13:00:00'), 'Jordan Lee', 'DCX', 'Automated weekly reporting workflow — saving 5 hours per week'],
    [new Date('2025-05-15T10:00:00'), 'David Chen', 'CDPS', 'Built the automated compliance checker — zero manual reviews needed'],
    [new Date('2025-04-22T11:00:00'), 'Nina Kowalski', 'CDPS', 'Created the self-serve analytics dashboard for cross-team metrics']
  ]);
  wins.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  wins.setFrozenRows(1);
  wins.autoResizeColumns(1, 4);

  // ─── DIRECTORY ───
  var dir = ss.insertSheet('Directory');
  dir.getRange(1, 1, 9, 10).setValues([
    ['Name', 'Team', 'Role', 'Location', 'Manager', 'BeverageType', 'Beverage', 'Weekend', 'HiddenTalent', 'PhotoURL'],
    ['Aisha Khan', 'DCX', 'CX Lead', 'London, GMT', 'Tom Rivera', 'Tea', 'Earl grey, no sugar', 'Weekend pottery classes', 'Can name any country\'s capital in under 3 seconds', ''],
    ['Jordan Lee', 'DCX', 'Sr UX Designer', 'San Francisco, PT', 'Tom Rivera', 'Coffee', 'Flat white, oat milk', 'Mountain biking in Marin', 'Speaks 4 languages fluently', ''],
    ['Raj Patel', 'DCX', 'CX Analyst', 'Remote, ET', 'Sarah Park', 'Coffee', 'Cold brew, black', 'Board game nights', 'Built a working arcade cabinet from scratch', ''],
    ['Sameer Rajwade', 'DCX', 'Sr Associate', 'Remote, ET', 'Tom Rivera', 'Coffee', 'Americano with oat milk', 'Cricket and hiking', '', ''],
    ['Priya Sharma', 'CDPS', 'Sr Data Engineer', 'New York, ET', 'David Chen', 'Coffee', 'Oat latte, extra shot', 'Trail running in the Catskills', 'Solves Rubik\'s cubes in under 2 minutes', ''],
    ['Nina Kowalski', 'CDPS', 'Data Analyst', 'Chicago, CT', 'David Chen', 'Tea', 'Matcha latte', 'Baking sourdough', 'Won a pie baking contest 3 times', ''],
    ['Marcus Thompson', 'CAMS', 'Campaign Manager', 'Chicago, CT', 'Leila Hassan', 'Coffee', 'Black, no sugar, strong', 'Competitive chess tournaments', 'Can play 4 musical instruments', ''],
    ['Emily Rodriguez', 'CAMS', 'Marketing Analyst', 'Dallas, CT', 'Leila Hassan', 'Tea', 'Chai latte, extra spicy', 'Watercolor painting', 'Has visited 30 countries before age 30', '']
  ]);
  dir.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  dir.setFrozenRows(1);
  dir.autoResizeColumns(1, 10);

  // ─── CONFIG ───
  var config = ss.insertSheet('Config');
  config.getRange(1, 1, 6, 2).setValues([
    ['Key', 'Value'],
    ['PollQuestion', 'What should our September team activity be?'],
    ['PollOptions', 'Trivia Challenge,Speed Networking,Problem Jam'],
    ['BuzzQuestion', 'What\'s the worst movie you secretly love?'],
    ['STFormURL', 'https://forms.gle/your-st-form-url-here'],
    ['WinFormURL', 'https://forms.gle/your-win-form-url-here']
  ]);
  config.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  config.setFrozenRows(1);
  config.autoResizeColumns(1, 2);

  // ─── RATINGS ───
  var ratings = ss.insertSheet('Ratings');
  ratings.getRange(1, 1, 8, 3).setValues([
    ['Timestamp', 'EventTitle', 'Rating'],
    [new Date('2025-06-20T10:00:00'), 'June Pilot', 5],
    [new Date('2025-06-20T10:15:00'), 'June Pilot', 4],
    [new Date('2025-06-20T11:00:00'), 'June Pilot', 5],
    [new Date('2025-05-16T09:00:00'), 'May Planning Kickoff', 4],
    [new Date('2025-05-16T09:30:00'), 'May Planning Kickoff', 5],
    [new Date('2025-04-18T10:00:00'), 'April Tool Spotlight', 4],
    [new Date('2025-04-18T10:45:00'), 'April Tool Spotlight', 5]
  ]);
  ratings.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  ratings.setFrozenRows(1);
  ratings.autoResizeColumns(1, 3);

  // ─── BUZZ RESPONSES ───
  var buzz = ss.insertSheet('BuzzResponses');
  buzz.getRange(1, 1, 6, 4).setValues([
    ['Timestamp', 'Type', 'Question', 'Answer'],
    [new Date('2025-06-15T10:00:00'), 'Poll', 'What should our September team activity be?', 'Trivia Challenge'],
    [new Date('2025-06-15T10:30:00'), 'Poll', 'What should our September team activity be?', 'Speed Networking'],
    [new Date('2025-06-15T11:00:00'), 'Poll', 'What should our September team activity be?', 'Trivia Challenge'],
    [new Date('2025-06-16T09:00:00'), 'Question', 'What\'s the worst movie you secretly love?', 'The Room — I\'ve seen it 12 times'],
    [new Date('2025-06-16T14:00:00'), 'Question', 'What\'s the worst movie you secretly love?', 'Sharknado 3 and I\'m not ashamed']
  ]);
  buzz.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  buzz.setFrozenRows(1);
  buzz.autoResizeColumns(1, 4);

  // Final log
  Logger.log('─────────────────────────────────────');
  Logger.log('Three60 sheet created successfully!');
  Logger.log('Sheet ID: ' + ss.getId());
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('Copy the Sheet ID above into Code.gs');
  Logger.log('─────────────────────────────────────');

  SpreadsheetApp.setActiveSpreadsheet(ss);
}
