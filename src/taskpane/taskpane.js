/* eslint-disable no-undef */
import { welcomePromt, botUrl } from "../constants";
Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    displayChatMessage("", "Hi, I am your word assistant bot-NoviWord", "bot");
    // Handle the Ask button click
    document.getElementById("askButton").onclick = async function () {
      const question = document.getElementById("userInput").value;
      if (question) {
        document.getElementById("headerId").style.display = "none";
        //const response =
        //await
        initializeDirectLine(question);
        // displayChatMessage(question, response);
      }
    };

    document.getElementById("userInput").addEventListener("keydown", function (event) {
      if (event.key === "Enter") {
        // Check if Enter key is pressed
        event.preventDefault(); // Prevents the default behavior (like submitting a form)

        const question = document.getElementById("userInput").value;
        if (question) {
          initializeDirectLine(question);
        }
      }
    });

    // Handle the Insert button click
    document.getElementById("insertButton").onclick = async function () {
      const response = document.getElementById("chatWindow").lastChild
        ? document.getElementById("chatWindow").lastChild.innerText
        : "";
      if (response) {
        await insertResponseIntoDocument(response);
      }
    };
  }
});

window.onDocumentOpen = function (event) {
  console.log("Testing>>>>Document Open Event Triggered");

  // Try opening the task pane automatically
  Office.addin
    .showAsTaskpane()
    .then(() => console.log("Taskpane opened"))
    .catch((err) => console.error("Error opening taskpane:", err));

  event.completed(); // Notify Office that the event is handled
};

// Register the event handler
Office.actions.associate("onDocumentOpen", window.onDocumentOpen);
// Function to get the chatbot's response (simple hardcoded response or integrate with an API)
// async function getChatbotResponse(question) {
//   // For now, a simple mock response
//   // console.log("Testing");
//   // Example Usage:
//   // fetchGeminiResponse("Tell me a fun fact about space.");
//   initializeDirectLine(question);
//   // return "This is a response to: " + question;
// }

// Display user question and bot response in chat window
function displayChatMessage(question, response, role) {
  const chatWindow = document.getElementById("chatWindow");
  // if (role === "bot") {
  //   chatWindow.innerHTML += `<div class="bot"><img src="../../assets/copilot.png"/> <br>${response}</div>`;
  // } else {
  //   chatWindow.innerHTML += `<div class="user">You<br>${question}</div>`;
  // }
  if (role === "bot") {
    chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="../../assets/copilot.png"/> NoviWord</div><div class="message bot">${response}</div>`;
  } else {
    chatWindow.innerHTML += `<div class="user-wrapper">You</div><div class="message user">${question}</div>`;
  }
  document.getElementById("userInput").value = ""; // Clear input field
}

// Function to insert the response into the Word document
async function insertResponseIntoDocument(response) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertText(response, Word.InsertLocation.end);
    await context.sync();
  });
}
const initializeDirectLine = async function (question) {
  try {
    const response = await fetch(
      "https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent_gGhuq5/directline/token?api-version=2022-03-01-preview"
    );
    const data = await response.json();
    // console.log("Testing data token:" + JSON.stringify(data, null, 2));
    // console.log("DirectLine Object:", window.DirectLine);
    const directLine = new window.DirectLine.DirectLine({ token: data.token });
    // console.log("directLine*******", directLine);
    // console.log("DirectLine instance:", directLine);
    // console.log("DirectLine activity$:", directLine.activity$);

    if (!directLine || !directLine.activity$) {
      throw new Error("DirectLine instance failed to initialize");
    }
    directLine
      .postActivity({
        from: { id: "10", name: "User" },
        type: "message",
        text: question,
      })
      .subscribe(
        (id) => console.log("Message sent with ID:", id),
        (error) => console.error("Error sending message:", error)
      );

    directLine.activity$.subscribe((activity) => {
      console.log(welcomePromt);
      console.log(botUrl);
      console.log("Testing activity: ", activity);
      console.log("Role*******", activity.from.role);
      if (activity.type === "message" && activity.from.id !== "10" && !activity.recipient) {
        console.log("Testing response: ", activity.text);
        displayChatMessage(question, activity.text, activity.from.role);
      }
    });
  } catch (error) {
    console.error("Error initializing DirectLine:", error);
  }
};
