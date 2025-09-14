import * as React from "react";
import { HttpClient, HttpClientResponse } from "@microsoft/sp-http";

interface IFlowTriggerProps {
  context: any; // pass the web part context from your main component
}

const FlowTrigger: React.FC<IFlowTriggerProps> = ({ context }) => {
  const [responseMessage, setResponseMessage] = React.useState<string>("");

  const triggerFlow = async () => {
    const powerautomateURL =
      "https://81b07adac380e965b84fe5494a9635.dd.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/fd094997c0ff42109252ca4adcdda243/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_TyMnzAMImPsoe2bg-qdrRtMw_HV3jT-DyjkHyGg-UM"; // Replace with your flow's URL

    try {
      console.log("Triggering Power Automate...");

      const response: HttpClientResponse = await context.httpClient.post(
        powerautomateURL,
        HttpClient.configurations.v1,
        {
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ message: "Triggered from SPFx React!" }),
        }
      );

      if (response.ok) {
        const data = await response.json();
        console.log("Power Automate triggered successfully:", data);
        setResponseMessage("✅ Flow triggered successfully!\n List Item created ID: " + data.ID);
      } else {
        const errorText = await response.text();
        console.error("Failed to trigger Power Automate:", response.statusText, errorText);
        setResponseMessage("❌ Failed to trigger flow: " + response.statusText);
      }
    } catch (error) {
      console.error("Error triggering Power Automate:", error);
      setResponseMessage("⚠️ Error: " + error);
    }
  };

  return (
    <div>
      <button
        onClick={triggerFlow}
        style={{
          padding: "8px 16px",
          background: "#0078d4",
          color: "#fff",
          border: "none",
          borderRadius: "4px",
          cursor: "pointer",
        }}
      >
        Trigger Power Automate
      </button>
      <div style={{ marginTop: "10px" }}>{responseMessage}</div>
    </div>
  );
};

export default FlowTrigger;
