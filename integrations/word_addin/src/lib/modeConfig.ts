export enum ConnectorMode {
  Local = "local_open_source",
  Free = "free_connected",
  Enterprise = "enterprise_connected"
}

export const MODE_CONFIG = {
  [ConnectorMode.Local]: {
    buttonText: "Create Local GalloDoc",
    description: "Local mode: No data leaves your machine. Generates a local JSON.",
    requiresLogin: false
  },
  [ConnectorMode.Free]: {
    buttonText: "Save to HaloBridge",
    description: "Free mode: Basic cloud record and versioning. Uploads to HaloBridge.",
    requiresLogin: true
  },
  [ConnectorMode.Enterprise]: {
    buttonText: "Save to HaloBridge (Full)",
    description: "Enterprise mode: Fullgovernance, VerifyIQ & HIM-C review pipeline.",
    requiresLogin: true
  }
};
