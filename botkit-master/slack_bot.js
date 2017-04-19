/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
           ______     ______     ______   __  __     __     ______
          /\  == \   /\  __ \   /\__  _\ /\ \/ /    /\ \   /\__  _\
          \ \  __<   \ \ \/\ \  \/_/\ \/ \ \  _"-.  \ \ \  \/_/\ \/
           \ \_____\  \ \_____\    \ \_\  \ \_\ \_\  \ \_\    \ \_\
            \/_____/   \/_____/     \/_/   \/_/\/_/   \/_/     \/_/


This is a sample Slack bot built with Botkit.

This bot demonstrates many of the core features of Botkit:

* Connect to Slack using the real time API
* Receive messages based on "spoken" patterns
* Reply to messages
* Use the conversation system to ask questions
* Use the built in storage system to store and retrieve information
  for a user.

# RUN THE BOT:

  Get a Bot token from Slack:

    -> http://my.slack.com/services/new/bot

  Run your bot from the command line:

    token=<MY TOKEN> node slack_bot.js

# USE THE BOT:

  Find your bot inside Slack to send it a direct message.

  Say: "Hello"

  The bot will reply "Hello!"

  Say: "who are you?"

  The bot will tell you its name, where it is running, and for how long.

  Say: "Call me <nickname>"

  Tell the bot your nickname. NoZ//////w you are friends.

  Say: "who am I?"

  The bot will tell you your nickname, if it knows one for you.

  Say: "shutdown"

  The bot will ask if you are sure, and then shut itself down.

  Make sure to invite your bot into other channels using /invite @<my bot>!

# EXTEND THE BOT:

  Botkit has many features for building cool and useful bots!

  Read all about it here:

    -> http://howdy.ai/botkit

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/


if (!process.env.token) {
    console.log('Error: Specify token in environment');
    process.exit(1);
}

var Botkit = require('./lib/Botkit.js');
var os = require('os');
if (typeof require !== 'undefined') XLSX = require('./excelConverter/xlsx.js');
var workbook = XLSX.readFile('./excelConverter/obj.xlsx');
var fs = require('fs');
var controller = Botkit.slackbot({
    debug: true,
});

var bot = controller.spawn({
    token: process.env.token
}).startRTM();

const mongodb = require('mongodb')
const moment = require('moment')

// DB URLs
const devUrl = 'mongodb://mongodbtrigger:password@ds161580.mlab.com:61580/mongodbtrigger'
const prodUrl = 'mongodb://localhost:27017/history'

// ENV Vars
const url = process.env.NODE_ENV === "production" ? prodUrl : devUrl

controller.hears(['hello', 'hey'], 'direct_message', function (bot, message) {

    bot.api.reactions.add({
        timestamp: message.ts,
        channel: message.channel,
        name: 'robot_face',
    }, function (err, res) {
        if (err) {
            bot.botkit.log('Failed to add emoji reaction :(', err);
        }
    });


    controller.storage.users.get(message.user, function (err, user) {
        if (user && user.name) {
            bot.reply(message, 'Hello ' + user.name + '!!');
        } else {
            bot.reply(message, 'Hello.');
        }
    });
});

controller.hears(['call me (.*)', 'my name is (.*)'], 'direct_message,direct_mention,mention', function (bot, message) {
    var name = message.match[1];
    controller.storage.users.get(message.user, function (err, user) {
        if (!user) {
            user = {
                id: message.user,
            };
        }
        user.name = name;
        controller.storage.users.save(user, function (err, id) {
            bot.reply(message, 'Got it. I will call you ' + user.name + ' from now on.');
        });
    });
});

controller.hears(['what is my name', 'who am i'], 'direct_message,direct_mention,mention', function (bot, message) {

    controller.storage.users.get(message.user, function (err, user) {
        if (user && user.name) {
            bot.reply(message, 'Your name is ' + user.name);
        } else {
            bot.startConversation(message, function (err, convo) {
                if (!err) {
                    convo.say('I do not know your name yet!');
                    convo.ask('What should I call you?', function (response, convo) {
                        convo.ask('You want me to call you `' + response.text + '`?', [{
                                pattern: 'yes',
                                callback: function (response, convo) {
                                    // since no further messages are queued after this,
                                    // the conversation will end naturally with status == 'completed'
                                    convo.next();
                                }
                            },
                            {
                                pattern: 'no',
                                callback: function (response, convo) {
                                    // stop the conversation. this will cause it to end with status == 'stopped'
                                    convo.stop();
                                }
                            },
                            {
                                default: true,
                                callback: function (response, convo) {
                                    convo.repeat();
                                    convo.next();
                                }
                            }
                        ]);

                        convo.next();

                    }, {
                        'key': 'nickname'
                    }); // store the results in a field called nickname

                    convo.on('end', function (convo) {
                        if (convo.status == 'completed') {
                            bot.reply(message, 'OK! I will update my dossier...');

                            controller.storage.users.get(message.user, function (err, user) {
                                if (!user) {
                                    user = {
                                        id: message.user,
                                    };
                                }
                                user.name = convo.extractResponse('nickname');
                                controller.storage.users.save(user, function (err, id) {
                                    bot.reply(message, 'Got it. I will call you ' + user.name + ' from now on.');


                                });
                            });



                        } else {
                            // this happens if the conversation ended prematurely for some reason
                            bot.reply(message, 'OK, nevermind!');
                        }
                    });
                }
            });
        }
    });
});


controller.hears(['shutdown'], 'direct_message,direct_mention,mention', function (bot, message) {

    bot.startConversation(message, function (err, convo) {

        convo.ask('Are you sure you want me to shutdown?', [{
                pattern: bot.utterances.yes,
                callback: function (response, convo) {
                    convo.say('Bye!');
                    convo.next();
                    setTimeout(function () {
                        process.exit();
                    }, 3000);
                }
            },
            {
                pattern: bot.utterances.no,
                default: true,
                callback: function (response, convo) {
                    convo.say('*Phew!*');
                    convo.next();
                }
            }
        ]);
    });
});


controller.hears(['uptime2', 'identify yourself', 'who are you', 'what is your name'],
    'direct_message,direct_mention,mention',
    function (bot, message) {

        var hostname = os.hostname();
        var uptime = formatUptime(process.uptime());

        bot.reply(message,
            ':robot_face: I am a bot named <@' + bot.identity.name +
            '>. I have been running for ' + uptime + ' on ' + hostname + '.');

    });


controller.hears(['who is using (.*)', 'what is the status of (.*)'], 'direct_message,direct_mention', function (bot, message) {
    var completemachinenameinput = message.match[1];
    var machinenamechararray = completemachinenameinput.split('');
    var first_sheet_name = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[first_sheet_name];
    lengthofinput = completemachinenameinput.length - 1;
    var machinename = " ";
    var machinetype = " ";
    var machinelocation = " ";
    //to calculate machinename
    var numberofdash = 0;
    var lengthofinput1 = lengthofinput;
    //for number of dash calculation
    while (lengthofinput1 > 0) {
        if (machinenamechararray[lengthofinput1] == '-') {
            numberofdash++;
        }
        lengthofinput1--;
    }
    if (numberofdash == 2) {
        while ((machinenamechararray[lengthofinput] != '-')) //||(machinenamechararray[lengthofinput]!=" "))
        {
            machinename = machinename + completemachinenameinput[lengthofinput]
            lengthofinput--;

        }
        if (lengthofinput > 0) {
            lengthofinput--;
            //bot.reply(message, machinename);
            //to calculate machine type
            while ((machinenamechararray[lengthofinput] != '-')) //||(machinenamechararray[lengthofinput]!=" "))
            {
                machinetype = machinetype + completemachinenameinput[lengthofinput];
                lengthofinput--;
            }
        }
        if (lengthofinput > 0) {
            lengthofinput--;
            //bot.reply(message, machinename+machinetype+lengthofinput+);
            //to calculate machine location
            while (lengthofinput >= 0) {
                machinelocation = machinelocation + completemachinenameinput[lengthofinput];
                lengthofinput--;
            }
        }
    } else if (numberofdash == 1) {
        while ((machinenamechararray[lengthofinput] != '-')) //||(machinenamechararray[lengthofinput]!=" "))
        {
            machinename = machinename + completemachinenameinput[lengthofinput]
            lengthofinput--;

        }
        if (lengthofinput > 0) {
            lengthofinput--;
            //bot.reply(message, machinename);
            //to calculate machine type
            while (lengthofinput >= 0) //||(machinenamechararray[lengthofinput]!=" "))
            {
                machinetype = machinetype + completemachinenameinput[lengthofinput];
                lengthofinput--;
            }
        }

    } else {

        while (lengthofinput >= 0) //||(machinenamechararray[lengthofinput]!=" "))
        {
            machinename = machinename + completemachinenameinput[lengthofinput]
            lengthofinput--;

        }
    }

    var machinename = reverse(machinename);
    var machinetype = reverse(machinetype);
    var machinelocation = reverse(machinelocation);
    // bot.reply(message, 'machinename:-' +machinename +'machinetype:-' +machinelocation+ 'machinetype:-' +machinetype);
    //length of machine name should be of length 2 
    if ((machinename.length) == 2) {
        machinename = "0" + machinename;
    }
    var machinetypechararray = machinetype.split('');
    var lengthofmachinetype = machinetype.length;
    var c = machinetypechararray[0];
    var newmachinetype = " ";
    if (c >= '0' && c <= '9') {
        newmachinetype = newmachinetype + "one";
        for (var i = 1; i < lengthofmachinetype; ++i) { // it is a number
            newmachinetype = newmachinetype + machinetypechararray[i];
        }
    } else {
        newmachinetype = " " + machinetype;
    }
    if (machinelocation.length == 1) {
        machinelocation = " noi";
    }
    //removing spaces globally
    machinelocation = machinelocation.replace(/\s/g, "");
    machinename = machinename.replace(/\s/g, "");
    newmachinetype = newmachinetype.replace(/\s/g, "");
    //bot.reply(message,machinelocation+'-'+newmachinetype+'-'+machinename);

    // creating complete machine name
    var completeoneboxname = machinelocation + '-' + newmachinetype + '-' + machinename;
    // bot.reply(message,completeoneboxname);

    const mongoClient = mongodb.MongoClient

    mongoClient.connect(url, async(err, db) => {
        if (err) console.log(err)
        db.listCollections().toArray(function (err, collections) {
            const collection = collections.find(function (collection) {
                completeoneboxname === collection
            })

            if (!collection) console.log() //TODO: Reply... No entry for this 1box name found...
            else {
                db.collection(collection).findOne(function (err, host) {
                    if (host.login && host.logout) {
                        if (moment(host.logout).isAfter(host.login)) {
                            // TODO: no one is using vm... Last assigned to host.user
                        } else if (!moment(host.logout).isAfter(host.login)) {
                            // TODO: someone is using vm... Assigned to host.user
                        }
                    } else {
                        // Reply... login or logout script never ran on vm
                    }
                })
            }
            // find name completeoneboxname === hostmae
        })

        if (!!host) {
            console.log("Found entry for current host... Updating!")
            host.user = user
            host[taskType] = time
            await hosts.save(host)
        } else {
            console.log("No entry for current host found... Creating!")
            const query = {}
            query["user"] = user
            query[taskType] = time
            await hosts.insert(query)
        }
        db.close()
    })

    // var range = XLSX.utils.decode_range(worksheet['!ref']);
    // var flag = 0; // get the range
    // for (var R = range.s.r; R <= range.e.r; ++R) {
    //     for (var C = range.s.c; C <= range.e.c; ++C) {
    //         /* find the cell object */
    //         var cellref = XLSX.utils.encode_cell({
    //             c: C,
    //             r: R
    //         }); // construct A1 reference for cell
    //         if (!worksheet[cellref]) continue; // if cell doesn't exist, move on
    //         var cell = worksheet[cellref];

    //         /* if the cell is a text cell with the old string, change it */

    //         if (cell.v === completeoneboxname) {
    //             flag = 1;
    //             var cellref2 = XLSX.utils.encode_cell({
    //                 c: C + 1,
    //                 r: R
    //             }); // construct A1 reference for cell
    //             if (!worksheet[cellref2]) continue; // if cell doesn't exist, move on
    //             var cell2 = worksheet[cellref2];
    //             if (cell2.v === 'NA') {
    //                 bot.startConversation(message, function (err, convo) {
    //                     if (!err) {
    //                         convo.say('Great! this machine is not assigned to anyone till now');
    //                         convo.say('Please provide your name so that this machine will be assigned to you');
    //                         convo.ask('What should I call you?', function (response, convo) {
    //                             convo.ask('You want me to assign ' + completeoneboxname + 'to ' + response.text + '`?', [{
    //                                     pattern: 'yes',
    //                                     callback: function (response, convo) {
    //                                         // since no further messages are queued after this,
    //                                         // the conversation will end naturasslly with status == 'completed'
    //                                         convo.next();
    //                                     }
    //                                 },
    //                                 {
    //                                     pattern: 'no',
    //                                     callback: function (response, convo) {
    //                                         // stop the conversation. this will cause it to end with status == 'stopped'
    //                                         convo.stop();
    //                                     }
    //                                 },
    //                                 {
    //                                     default: true,
    //                                     callback: function (response, convo) {
    //                                         convo.repeat();
    //                                         convo.next();
    //                                     }
    //                                 }
    //                             ]);

    //                             convo.next();

    //                         }, {
    //                             'key': 'nickname'
    //                         }); // store the results in a field called nickname

    //                         convo.on('end', function (convo) {
    //                             if (convo.status == 'completed') {
    //                                 bot.reply(message, 'Done it');

    //                                 controller.storage.users.get(message.user, function (err, user) {
    //                                     if (!user) {
    //                                         user = {
    //                                             id: message.user,
    //                                         };
    //                                     }
    //                                     user.name = convo.extractResponse('nickname');
    //                                     controller.storage.users.save(user, function (err, id) {
    //                                         cell2.v = user.name;
    //                                         XLSX.writeFile(workbook, './excelConverter/obj.xlsx');
    //                                         bot.reply(message, 'Got it. The machine' + completeoneboxname + ' is assigned to ' + user.name + ' from now onwards.');


    //                                     });
    //                                 });



    //                             } else {
    //                                 // this happens if the conversation ended prematurely for some reason
    //                                 bot.reply(message, 'OK, nevermind!');
    //                             }
    //                         });
    //                     }
    //                 });

    //             } else {
    //                 bot.reply(message, 'This one Box is assigned to the user :-  ' + cell2.v);

    //             }

    //         }

    //     }
    // }
    // if (flag === 0) {
    //     bot.reply(message, 'There is no such virtual machine exist in my dossier please add it using my add machine functionality');
    // }
});
controller.hears(['add new machine (.*)', 'add a new machine (.*)'], 'direct_message,direct_mention', function (bot, message) {
    var completemachinenameinput = message.match[1];
    var machinenamechararray = completemachinenameinput.split('');
    var first_sheet_name = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[first_sheet_name];
    lengthofinput = completemachinenameinput.length - 1;
    var machinename = " ";
    var machinetype = " ";
    var machinelocation = " ";
    //to calculate machinename
    var numberofdash = 0;
    var lengthofinput1 = lengthofinput;
    //for number of dash calculation
    while (lengthofinput1 > 0) {
        if (machinenamechararray[lengthofinput1] == '-') {
            numberofdash++;
        }
        lengthofinput1--;
    }
    if (numberofdash == 2) {
        while ((machinenamechararray[lengthofinput] != '-')) //||(machinenamechararray[lengthofinput]!=" "))
        {
            machinename = machinename + completemachinenameinput[lengthofinput]
            lengthofinput--;

        }
        if (lengthofinput > 0) {
            lengthofinput--;
            // bot.reply(message, machinename);
            //to calculate machine type
            while ((machinenamechararray[lengthofinput] != '-')) //||(machinenamechararray[lengthofinput]!=" "))
            {
                machinetype = machinetype + completemachinenameinput[lengthofinput];
                lengthofinput--;
            }
        }
        if (lengthofinput > 0) {
            lengthofinput--;
            //bot.reply(message, machinename+machinetype+lengthofinput+);
            //to calculate machine location
            while (lengthofinput >= 0) {
                machinelocation = machinelocation + completemachinenameinput[lengthofinput];
                lengthofinput--;
            }
        }
        var machinename = reverse(machinename);
        var machinetype = reverse(machinetype);
        var machinelocation = reverse(machinelocation);
        // bot.reply(message, 'machinename:-' +machinename +'machinetype:-' +machinelocation+ 'machinetype:-' +machinetype);
        //length of machine name should be of length 2 
        if ((machinename.length) == 2) {
            machinename = "0" + machinename;
        }
        var machinetypechararray = machinetype.split('');
        var lengthofmachinetype = machinetype.length;
        var c = machinetypechararray[0];
        var newmachinetype = " ";
        if (c >= '0' && c <= '9') {
            newmachinetype = newmachinetype + "one";
            for (var i = 1; i < lengthofmachinetype; ++i) { // it is a number
                newmachinetype = newmachinetype + machinetypechararray[i];
            }
        } else {
            newmachinetype = " " + machinetype;
        }
        if (machinelocation.length == 1) {
            machinelocation = " noi";
        }
        //removing spaces globally
        machinelocation = machinelocation.replace(/\s/g, "");
        machinename = machinename.replace(/\s/g, "");
        newmachinetype = newmachinetype.replace(/\s/g, "");
        //bot.reply(message,machinelocation+'-'+newmachinetype+'-'+machinename);

        // creating complete machine name
        var completeoneboxname = machinelocation + '-' + newmachinetype + '-' + machinename;
        //bot.reply(message,completeoneboxname);
        var range = XLSX.utils.decode_range(worksheet['!ref']);
        var flag = 0; // get the range
        for (var R = range.s.r; R <= range.e.r; ++R) {
            for (var C = range.s.c; C <= range.e.c; ++C) {
                /* find the cell object */
                var cellref = XLSX.utils.encode_cell({
                    c: C,
                    r: R
                }); // construct A1 reference for cell
                if (!worksheet[cellref]) continue; // if cell doesn't exist, move on
                var cell = worksheet[cellref];

                /* if the cell is a text cell with the old string, change it */
                if (cell.v === completeoneboxname) {
                    flag = 1;
                }
            }
        }
        //bot.reply(message,'flag is'+flag)
        if (flag == 0) {
            var range = XLSX.utils.decode_range(worksheet['!ref']);
            for (var R = range.s.r; R <= range.e.r; ++R) {
                for (var C = range.s.c; C <= range.e.c; ++C) {
                    /* find the cell object */
                    var cellref = XLSX.utils.encode_cell({
                        c: C,
                        r: R
                    }); // construct A1 reference for cell
                    if (!worksheet[cellref]) continue; // if cell doesn't exist, move on
                    var cell = worksheet[cellref];
                    //bot.reply(message,'this is '+R+' '+C);
                    /* if the cell is a text cell with the old string, change it */
                    if (cell.v === "Blank") { //bot.reply(message,'this is '+R+' '+C);
                        cell.v = completeoneboxname;
                        XLSX.writeFile(workbook, './excelConverter/obj.xlsx');
                        bot.reply(message, 'Got it. The machine ' + completeoneboxname + ' is added successfully ');
                        R = range.e.r;
                        C = range.s.c;
                        /* var cell2 = { v: "NA" };
                                                                var cellref2 = XLSX.utils.encode_cell({c:C+1, r:R}); // construct A1 reference for cell
                                                                cell2.t = 's'
                                                                 // if cell doesn't exist, move on
                                                                worksheet[cellref2]=cell2;
                                                                worksheet['!ref'] = XLSX.utils.encode_range(range);
                                                                workbook.SheetNames.push(first_sheet_name);s
                                                                workbook.Sheets[first_sheet_name] = worksheet;
                                                                XLSX.writeFile(workbook, './excelConverter/obj.xlsx');
                                                                bot.reply(message, 'Got it. The machine'+completeoneboxname+ ' is added successfully ');

                                                                //var cellref3 = XLSX.utils.encode_cell({c:C, r:R+1}); // construct A1 reference for cell
                                                                //if(!worksheet[cellref3]) continue; // if cell doesn't exist, move on
                                                                //var cell3 = worksheet[cellref3];
                                                                //cell3.v="NEW";
                                                                //XLSX.writeFile(workbook, './excelConverter/obj.xlsx');
                                                                //sbot.reply(message, 'Got it. The machine'+completeoneboxname+ ' is added successfully ');
                                                            */
                    }
                }
            }

        }
        if (flag == 1) {
            bot.reply(message, ' This machine is already in my dossier please try with different machine');
        }
    } else {
        bot.reply(message, ' Please provide machine name in correct format i.e. machinelocation-machinetype-machinename');
    }
});

controller.hears(['remove machine (.*)', 'remove a machine (.*)'], 'direct_message,direct_mention', function (bot, message) {
    var completemachinenameinput = message.match[1];
    var machinenamechararray = completemachinenameinput.split('');
    var first_sheet_name = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[first_sheet_name];
    lengthofinput = completemachinenameinput.length - 1;
    var machinename = " ";
    var machinetype = " ";
    var machinelocation = " ";
    //to calculate machinename
    var numberofdash = 0;
    var lengthofinput1 = lengthofinput;
    //for number of dash calculation
    while (lengthofinput1 > 0) {
        if (machinenamechararray[lengthofinput1] == '-') {
            numberofdash++;
        }
        lengthofinput1--;
    }
    if (numberofdash == 2) {
        while ((machinenamechararray[lengthofinput] != '-')) //||(machinenamechararray[lengthofinput]!=" "))
        {
            machinename = machinename + completemachinenameinput[lengthofinput]
            lengthofinput--;

        }
        if (lengthofinput > 0) {
            lengthofinput--;
            // bot.reply(message, machinename);
            //to calculate machine type
            while ((machinenamechararray[lengthofinput] != '-')) //||(machinenamechararray[lengthofinput]!=" "))
            {
                machinetype = machinetype + completemachinenameinput[lengthofinput];
                lengthofinput--;
            }
        }
        if (lengthofinput > 0) {
            lengthofinput--;
            //bot.reply(message, machinename+machinetype+lengthofinput+);
            //to calculate machine location
            while (lengthofinput >= 0) {
                machinelocation = machinelocation + completemachinenameinput[lengthofinput];
                lengthofinput--;
            }
        }
        var machinename = reverse(machinename);
        var machinetype = reverse(machinetype);
        var machinelocation = reverse(machinelocation);
        // bot.reply(message, 'machinename:-' +machinename +'machinetype:-' +machinelocation+ 'machinetype:-' +machinetype);
        //length of machine name should be of length 2 
        if ((machinename.length) == 2) {
            machinename = "0" + machinename;
        }
        var machinetypechararray = machinetype.split('');
        var lengthofmachinetype = machinetype.length;
        var c = machinetypechararray[0];
        var newmachinetype = " ";
        if (c >= '0' && c <= '9') {
            newmachinetype = newmachinetype + "one";
            for (var i = 1; i < lengthofmachinetype; ++i) { // it is a number
                newmachinetype = newmachinetype + machinetypechararray[i];
            }
        } else {
            newmachinetype = " " + machinetype;
        }
        if (machinelocation.length == 1) {
            machinelocation = " noi";
        }
        //removing spaces globally
        machinelocation = machinelocation.replace(/\s/g, "");
        machinename = machinename.replace(/\s/g, "");
        newmachinetype = newmachinetype.replace(/\s/g, "");
        //bot.reply(message,machinelocation+'-'+newmachinetype+'-'+machinename);

        // creating complete machine name
        var completeoneboxname = machinelocation + '-' + newmachinetype + '-' + machinename;
        //bot.reply(message,completeoneboxname);
        var range = XLSX.utils.decode_range(worksheet['!ref']);
        var flag = 0; // get the range
        for (var R = range.s.r; R <= range.e.r; ++R) {
            for (var C = range.s.c; C <= range.e.c; ++C) {
                /* find the cell object */
                var cellref = XLSX.utils.encode_cell({
                    c: C,
                    r: R
                }); // construct A1 reference for cell
                if (!worksheet[cellref]) continue; // if cell doesn't exist, move on
                var cell = worksheet[cellref];

                /* if the cell is a text cell with the old string, change it */
                if (cell.v === completeoneboxname) {
                    flag = 1;
                    cell.v = "Deleted";
                    var cellref3 = XLSX.utils.encode_cell({
                        c: C + 1,
                        r: R
                    }); // construct A1 reference for cell
                    if (!worksheet[cellref3]) continue; // if cell doesn't exist, move on
                    var cell3 = worksheet[cellref3];
                    cell3.v = "Deleted";
                    XLSX.writeFile(workbook, './excelConverter/obj.xlsx');
                    bot.reply(message, 'Got it. The machine' + completeoneboxname + ' is deleted successfully ');

                }
            }
        }
        //bot.reply(message,'flag is'+flag)
        if (flag == 0) {
            bot.reply(message, ' This machine is not in my dossier please try with different machine'); //sbot.reply(message, 'Got it. The machine'+completeoneboxname+ ' is added successfully ');
        }
    } else {
        bot.reply(message, ' Please provide machine name in correct format i.e. machinelocation-machinetype-machinename');
    }
});

function formatUptime(uptime) {
    var unit = 'second';
    if (uptime > 60) {
        uptime = uptime / 60;
        unit = 'minute';
    }
    if (uptime > 60) {
        uptime = uptime / 60;
        unit = 'hour';
    }
    if (uptime != 1) {
        unit = unit + 's';
    }

    uptime = uptime + ' ' + unit;
    return uptime;
}

function reverse(s) {
    return s.split("").reverse().join("");
}
