/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */
 
Office.onReady(async function (info) {
  displayStartingMessage("Hi, I am your word assistant bot-NoviWord");
  let directLine1 = await initializeDirectLine();
if (info.host === Office.HostType.Word) {
  //let flag=true;
 
document.getElementById("askButton").onclick = async function () {
  const question = document.getElementById("userInput").value;
  if (question) {
    document.getElementById("headerId").style.display = "none"; 
    displayChatMessage(question, '', "User");
      await getBotResponse(directLine1, question);
   
 
  }
};
 
document.getElementById("userInput").addEventListener("keydown", async function (event) {
  if (event.key === "Enter") {
    // Check if Enter key is pressed
    event.preventDefault(); // Prevents the default behavior (like submitting a form)
    document.getElementById("headerId").style.display = "none";
    const question = document.getElementById("userInput").value;
    if (question) {
     
        displayChatMessage(question, '', "User");
      await getBotResponse(directLine1, question);
     
  }
}});
 
// Handle the Insert button click
document.getElementById("insertButton").onclick = async function () {
  const response = document.getElementById("chatWindow").lastChild
    ? document.getElementById("chatWindow").lastChild.innerHTML
    : "";
  if (response) {
    await insertResponseIntoDocument(response);
  }
};
}
});


function displayStartingMessage(starter) {
  const chatWindow = document.getElementById("chatWindow");
  chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">${starter}</div>`;      
}
// Display user question and bot response in chat window
function displayChatMessage(question, response, role) {
  const chatWindow = document.getElementById("chatWindow");
 
  // Check if response is valid and if attachments exist
  if (response && response.attachments && response.attachments.length > 0) {
    response.attachments.forEach((attachment) => {
      // Check if attachment content has 'buttons' and 'signin' type
      if (attachment.content && attachment.content.buttons && attachment.content.buttons.length > 0) {
        attachment.content.buttons.forEach((button) => {
          if (button.type === "signin") {
            // Create a sign-in button
            const signinButton = document.createElement("button");
            signinButton.innerText = button.title || "Sign In"; // Default title to "Sign In"
            signinButton.classList.add("ms-Button", "ms-Button--primary");
 
            // Open the sign-in URL when the button is clicked
            signinButton.onclick = () => {
              window.open(button.value, "_blank"); // Open the sign-in URL in a new tab
            };
 
            // Display the bot's message
            chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">${attachment.content.text}</div>`;
            chatWindow.appendChild(signinButton); // Add the button after the message
          }
        });
      }
    });
  } else {
    // Regular message display if no attachments
    if (role === "bot") {
      if(response.speak==="Generate"){
        insertResponseIntoDocument(response.text);
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">SOW content generated in document</div>`;
      }
      else if(response.text){
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">${response.text}</div>`;      }
    } else {
      if(question){
     
        chatWindow.innerHTML += `<div class="user-wrapper">You</div><div class="message user">${question}</div>`;      }
     
    }
  }
  scrollToBottom();
  // Clear the input field
  document.getElementById("userInput").value = "";
}
 
// Function to insert the response into the Word document
async function insertResponseIntoDocument(response) {
  console.log("Testing insert to doc*********");
  console.log("response*******", response);
  await Word.run(async (context) => {
    console.log("Inside Testing insert to doc*********");
  console.log("Inside response*******", response);
    const body = context.document.body;
    body.insertHtml(response, Word.InsertLocation.end);
    await context.sync();
  });
}
const initializeDirectLine = async function () {
  try {
    const response = await fetch(
      "https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview"
    );
    const data = await response.json();
   
    const directLine = new window.DirectLine.DirectLine({ token: data.token });
   
 
    if (!directLine || !directLine.activity$) {
      throw new Error("DirectLine instance failed to initialize");
    }
 
    directLine
      .postActivity({
        from: { id: "10", name: "User" },
        type: "message",
        text: "Hi",
      })
      .subscribe(
        (id) => console.log("Message sent with ID:", id),
        (error) => console.error("Error sending message:", error)
      );
 
    directLine.activity$.subscribe((activity) => {
      console.log("Testing activity: ", activity);
      console.log("Role", activity.from.role);
      if (activity.type === "message" && activity.from.id !== "10" && !activity.recipient) {
        console.log("Bot Response: ", activity.text);
        displayChatMessage(false, activity, activity.from.role);
       
      }
    });
    return directLine;
  } catch (error) {
    console.error("Error initializing DirectLine:", error);
  }
};
 
const getBotResponse = async function (directLine, question) {
  directLine
    .postActivity({
      from: { id: "10", name: "User" },
      type: "message",
      text: question,
    })
    .subscribe(             // calls the subscription already created in InitializeDirectline method
      (id) => console.log("Message sent with ID:", id),
      (error) => console.error("Error sending message:", error)
    );
 
}
function scrollToBottom() {
  const chatWindow = document.getElementById("chatWindow");
  setTimeout(() => {
    chatWindow.scrollTop = chatWindow.scrollHeight;
  }, 100); // Timeout ensures scroll happens after the new message is rendered
}

let recognition = null;

document.getElementById("speakButton").addEventListener("click", async () => {
  try {
    await navigator.mediaDevices.getUserMedia({ audio: true });
    console.log("Microphone access granted.");
    startVoiceInput();
  } catch (error) {
    console.error("Microphone access denied:", error);
  }
});

function startVoiceInput() {
  // If a recognition instance already exists, stop it before starting a new one
  if (recognition && recognition.abort) {
    recognition.abort(); // This will stop the ongoing recognition session
  }

  recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
  recognition.lang = "en-US";
  recognition.interimResults = false;
  recognition.maxAlternatives = 1;
  recognition.continuous = true;

  // recognition.start();
   // trying continuos
  recognition.onstart = function () {
    console.log("Speech recognition started.");
  };
   // trying continuos
  recognition.onresult = function (event) {
    const transcript = event.results[0][0].transcript;
    console.log("Recognized text:", transcript);
    // insertTextIntoWord(transcript);
    document.getElementById("userInput").value = transcript;
  };

  recognition.onerror = function (event) {
    console.error("Speech recognition error:", event.error);
     // trying continuos
    if (event.error === 'aborted') {
      console.log("Restarting recognition...");
      recognition.start();
    }
     // trying continuos
  };
  // trying continuos
  recognition.onend = function () {
    console.log("Speech recognition ended.");
  };

  recognition.start();
  //  // trying continuos
}
