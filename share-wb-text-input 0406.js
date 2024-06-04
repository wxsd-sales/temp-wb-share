
/********************************************************
* 
* Macro Author:  Victor Vazquez
*                Technical Solutions Architect
*                vvazquez@cisco.com
*                Cisco Systems
* 
* Version: 1-0-0
* Released: 12/04/24
* 
* Version: 2-0-0
* Released: 03/06/24
* 
* This Webex Device macro allows users to share whiteboards
* via email simply by clicking on a button on the Navigator. 
*
* This specific version of this macro has been designed for 
* Companion mode use case. This macro should be installed and 
* enabled on the main Room Device. The macro then lets a user
* share a Whiteboard which may be open on the Companion Board
* from the main Room Devices Navigator.
* 
* Full Readme and source code and license details available here:
* https://github.com/wxsd-sales/share-whiteboard-macro
* 
********************************************************/

import xapi from 'xapi';

/*********************************************************
 * Configure the settings below
**********************************************************/

let emailConfig = {
  destination: 'vvazquez@cisco.com', // Change this value to the email address you want the whiteboard to be sent to by default
  body: 'Here you have your white board', // Email body text of your choice, this is an example
  subject: 'New white board', // Email subject of your choice, this is an example
  attachmentFilename: 'myfile-companion-mode' // File name of your choice, this is an example
};
const buttonConfig = {
  name: 'Send whiteboard',
  icon: 'Tv',
  panelId: 'share-wb'
};

const remoteDeviceconfig = {
  deviceIP: '192.168.100.150', // Change this value to the Webex Board IP address
  userName: 'victor', // Username with admin access to the Webex Board
  password: 'cisco,123', // Password for user above
};

const credentials = btoa(`${remoteDeviceconfig.userName}:${remoteDeviceconfig.password}`);

let boardUrl = '';
/*********************************************************
 * Main functions and event subscriptions
 **********************************************************/

// Create UI Extension Panel
createPanel();

// Set HTTP Client Config and 
xapi.Config.HttpClient.Mode.set('On');
xapi.Config.HttpClient.AllowInsecureHTTPS.set('True');


// Listen for Panel Clicks
xapi.Event.UserInterface.Extensions.Panel.Clicked.on(shareWhiteBoard);

// Process customer answer to Input Dialog Box for email destination
xapi.Event.UserInterface.Message.TextInput.Response.on(processInput)

/*********************************************************
 * Instructs the Companion Device to send the Whitebard
 * to configured email destination
 **********************************************************/

function sendWhiteBoardUrl(url) {
  const xml = ` <Command>
                  <Whiteboard>
                    <Email>
                      <Send>
                        <Subject>${emailConfig.subject}</Subject>
                        <Body>${emailConfig.body}</Body>
                        <Recipients>${emailConfig.destination}</Recipients>
                        <BoardUrls>${url}</BoardUrls>
                        <AttachmentFilenames>${emailConfig.attachmentFilename}.pdf</AttachmentFilenames>
                      </Send>
                    </Email>
                  </Whiteboard>
                </Command>`;
  xapi.Command.HttpClient.Post({
    AllowInsecureHTTPS: 'True',
    Header: ['Authorization: Basic ' + credentials],
    Url: `https://${remoteDeviceconfig.deviceIP}/putxml`
  }, xml)
    .catch(error => console.log('Error sending whiteboard', error))
    .then(reponse => {
      console.log('putxml response status code:', reponse.StatusCode);
      alert({ message: `Whiteboard has been sent to ${emailConfig.destination}` });
    })
}

/*****************************************************************
 * Listen for Panel Click Events and check Whiteboard xStatus
 * of Companion Device before attempting to send the Whiteboard
 * Presents Text Input Dialog Box asking for the destination email
 *****************************************************************/
async function shareWhiteBoard(event) {
  if (event.PanelId != buttonConfig.panelId) return;

  console.log(`Button ${buttonConfig.panelId} clicked`);
  console.log('Checking Companion Board WhiteBoard status')

  alert({ message: 'Checking for visible whiteboards on Companion Board', duration: 5 });

  // Get Board URL from the WB
  await xapi.Command.HttpClient.Get({
    AllowInsecureHTTPS: 'True',
    Header: ['Authorization: Basic ' + credentials],
    Url: `https://${remoteDeviceconfig.deviceIP}/getxml?location=/Status/Conference/Presentation/WhiteBoard`
  })
    .then(response => {
      boardUrl = response.Body.split("<BoardUrl>")[1].split("</BoardUrl>")[0];
      if (!boardUrl) {
        alert({ title: 'Warning', message: 'You need to share a whiteboard before it can be sent' });
        return
      }
    })
    .catch(error => {
      console.log('Error getting Board URL;', error);
      console.log(boardUrl);
    })
  if (boardUrl == '') {
    alert('You need to share a White Board before it can be sent');
    return;
  }
  else {
    console.log('Getting the email address');
    await xapi.Command.UserInterface.Message.TextInput.Display(
      {
        Duration: 300,
        FeedbackId: 'text-input-box',
        InputText: emailConfig.destination,
        InputType: 'SingleLine',
        KeyboardState: 'Open',
        SubmitText: 'Send',
        Text: 'Type the email address you want to share the whiteboard with',
        Title: 'Sending Whiteboard'
      }
    )
      /* wait for Customer answer in another function */
      .catch(error => console.log('Error getting email address', error));
  }
}

/*********************************************************
 * Gets destination email address
 **********************************************************/

function processInput(event) {
  if (event.FeedbackId !== 'text-input-box') return;
  console.log('saving email address', event.Text);
  emailConfig.destination = event.Text;
  console.log('Instructing Webex Board to send whiteboard:', boardUrl);
  sendWhiteBoardUrl(boardUrl); // Instruct Board to send URL
}

/**
 * Alert Function for Logging & Displaying Notification on Device
 * @property {object}  args               - Alert details
 * @property {string}  args.message       - Message Text
 * @property {string}  args.title         - Alert Title
 * @property {number}  args.duration      - Alert Duration
 */
function alert(args) {

  if (!args.hasOwnProperty('message')) {
    console.error('message is required to display alert')
    return
  }

  let duration = 5;
  if (args.hasOwnProperty('duration')) {
    duration = args.duration;
  }

  console.log('Displaying Alert:', args)
  if (args.hasOwnProperty('title')) {
    switch (args.title.toLowerCase()) {
      case 'warning':
        xapi.Command.UserInterface.Message.Alert.Display({ Duration: 10, Text: args.message, Title: args.title });
        break;
      default:
        xapi.Command.UserInterface.Message.Prompt.Display({ Duration: duration, Text: args.message, Title: args.title });
    }
  } else {
    xapi.Command.UserInterface.Message.Prompt.Display({ Duration: duration, Text: args.message, Title: 'Sharing Whiteboard Macro' })
  }
}


/*********************************************************
 * Create the UI Extension Panel and Save it to the Device
 **********************************************************/
async function createPanel() {
  const panelId = buttonConfig.panelId;

  let order = "";
  const orderNum = await panelOrder(panelId);
  if (orderNum != -1) order = `<Order>${orderNum}</Order>`;

  const panel = `
    <Extensions>
      <Panel>
        ${order}
        <Origin>local</Origin>
        <Location>CallControls</Location>
        <Icon>${buttonConfig.icon}</Icon>
        <Name>${buttonConfig.name}</Name>
        <ActivityType>Custom</ActivityType>
      </Panel>
    </Extensions>`
  xapi.Command.UserInterface.Extensions.Panel.Save(
    { PanelId: panelId },
    panel
  )
    .catch(e => console.log('Error saving panel: ' + e.message))
}


/*********************************************************
 * Gets the current Panel Order if exiting Macro panel is present
 * to preserve the order in relation to other custom UI Extensions
 **********************************************************/
async function panelOrder(panelId) {
  const list = await xapi.Command.UserInterface.Extensions.List({
    ActivityType: "Custom",
  });
  if (!list.hasOwnProperty("Extensions")) return -1;
  if (!list.Extensions.hasOwnProperty("Panel")) return -1;
  if (list.Extensions.Panel.length == 0) return -1;
  for (let i = 0; i < list.Extensions.Panel.length; i++) {
    if (list.Extensions.Panel[i].PanelId == panelId)
      return list.Extensions.Panel[i].Order;
  }
  return -1;
}