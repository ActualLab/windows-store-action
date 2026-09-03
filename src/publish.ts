require("dotenv").config();
import * as core from "@actions/core";

import fs = require("fs");
import path = require("path");
import api = require("./common/apiHelper");
import request = require("./common/requestHelper");
import Q = require("q");

/** The current token used for authentication. */
var currentToken: request.AccessToken;

/** The app ID we are publishing to */
var appId: string;

var packages: string[] = [];

/** Expected imageType values */
const imageType: string[] = [
  "Screenshot",
  "MobileScreenshot",
  "XboxScreenshot",
  "SurfaceHubScreenshot",
  "HoloLensScreenshot",
  "StoreLogo9x16",
  "StoreLogoSquare",
  "Icon",
  "PromotionalArt16x9",
  "PromotionalArtwork2400X1200",
  "XboxBrandedKeyArt",
  "XboxTitledHeroArt",
  "XboxFeaturedPromotionalArt",
  "SquareIcon358X358",
  "BackgroundImage1000X800",
  "PromotionalArtwork414X180",
];

/** The following attributes are considered as lists of strings and not just strings. */
const STRING_ARRAY_ATTRIBUTES = {
  keywords: true,
  features: true,
  recommendedhardware: true,
};

const packageExtensions = [
  ".msix", ".msixbundle", ".msixupload", ".appx", ".appxbundle", ".appxupload", ".xap",
];

const MODES = ["publish", "upload", "commit"];
const POLL_UNTIL = ["certification", "published"];

/**
 * The main task function.
 *
 * mode=upload   creates a submission, uploads the packages and leaves it pending.
 * mode=commit   takes the pending submission, applies release notes and commits it.
 * mode=publish  does both in one go (the original behavior).
 */
export async function publishTask() {
  /* We expect the endpoint part of this to not have a slash at the end.
   * This is because authenticating to 'endpoint/' will give us an
   * invalid token, while authenticating to 'endpoint' will work */
  api.ROOT =
    "https://manage.devcenter.microsoft.com" + api.API_URL_VERSION_PART;

  var mode = core.getInput("mode") || "publish";
  if (MODES.indexOf(mode) < 0) {
    throw new Error(`Unknown mode '${mode}'; expected one of ${MODES.join(", ")}`);
  }
  var pollUntil = core.getInput("poll-until") || "published";
  if (POLL_UNTIL.indexOf(pollUntil) < 0) {
    throw new Error(`Unknown poll-until '${pollUntil}'; expected one of ${POLL_UNTIL.join(", ")}`);
  }

  var credentials = {
    tenant: core.getInput("tenant-id"),
    clientId: core.getInput("client-id"),
    clientSecret: core.getInput("client-secret"),
  };

  console.log("Authenticating...");
  currentToken = await request.authenticate(
    "https://manage.devcenter.microsoft.com",
    credentials
  );
  appId = core.getInput("app-id"); // Globally set app ID for future steps.

  var submissionResource =
    mode === "commit"
      ? await getPendingSubmission()
      : await createSubmissionWithPackages();

  var submissionUrl = `https://developer.microsoft.com/en-us/dashboard/apps/${appId}/submissions/${submissionResource.id}`;
  core.setOutput("submission-id", submissionResource.id);
  core.setOutput("url", submissionUrl);

  if (mode === "upload") {
    console.log(`Packages are staged in pending submission ${submissionUrl}`);
    console.log("Run this action with mode=commit to send it to certification");
    return;
  }

  var releaseNotesPath = core.getInput("release-notes-path");
  if (releaseNotesPath) {
    console.log(`Applying release notes from ${releaseNotesPath}...`);
    setReleaseNotes(submissionResource, fs.readFileSync(releaseNotesPath, "utf8").trim());
    await putSubmission(submissionResource);
  }

  console.log("Committing submission...");
  await commitAppSubmission(submissionResource.id);

  if (core.getInput("skip-polling") === "true") {
    console.log("Skip polling option is checked. Skipping polling...");
    console.log(
      `Click here ${submissionUrl} to check the status of the submission in Dev Center`
    );
  } else {
    console.log(`Polling submission until ${pollUntil}...`);
    var resourceLocation = `applications/${appId}/submissions/${submissionResource.id}`;
    await api.pollSubmissionStatus(
      currentToken,
      resourceLocation,
      submissionResource.targetPublishMode,
      pollUntil === "certification"
    );
  }

  console.log("Submission completed");
}

/**
 * Creates a new submission (replacing the pending one when allowed), attaches the packages
 * found in package-path and uploads them.
 * @return Promises the new submission resource.
 */
async function createSubmissionWithPackages(): Promise<any> {
  var packagePath = core.getInput("package-path");
  var files = fs.readdirSync(packagePath);

  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    for (var j = 0; j < packageExtensions.length; j++) {
      var ext = packageExtensions[j];
      if (path.extname(file) == ext) {
        packages.push(file);
      }
    }
  }

  packages = packages.map((file) => {
    return path.join(packagePath, file);
  });

  if (core.getInput("delete-pending") !== "false") {
    await deletePendingSubmission();
  }

  console.log("Creating submission...");
  var submissionResource = await createAppSubmission();
  console.log(`Submission ${submissionResource.id} was created successfully`);

  if (core.getInput("delete-packages") === "true") {
    console.log("Deleting old packages...");
    api.deleteOldPackages(
      submissionResource.applicationPackages,
      +core.getInput("packages-keep")
    );
  }

  console.log("Updating submission...");
  await putMetadata(submissionResource);

  console.log("Creating zip file...");

  var zip = api.createZipFromPackages(packages);
  addImagesToZip(submissionResource, zip);

  // There might be no files in the zip if the user didn't supply any packages or images.
  // If there are files, persist the file.
  if (Object.keys(zip.files).length > 0) {
    await api.persistZip(zip, "temp.zip", submissionResource.fileUploadUrl);
  }

  return submissionResource;
}

/**
 * @return Promises the pending submission of the app; fails when there is none.
 */
async function getPendingSubmission(): Promise<any> {
  var app = await getAppResource();
  var pending = app.pendingApplicationSubmission;
  if (!pending || !pending.id) {
    throw new Error(
      `App ${appId} has no pending submission; run this action with mode=upload first`
    );
  }
  console.log(`Found pending submission ${pending.id}`);
  return api.getSubmission(
    currentToken,
    api.ROOT + "applications/" + appId + "/submissions/" + pending.id
  );
}

/**
 * Deletes the app's pending submission, if any. Submissions created in the Dev Center UI
 * can't be deleted through the API; that error surfaces from the delete call.
 */
async function deletePendingSubmission(): Promise<void> {
  var app = await getAppResource();
  var pending = app.pendingApplicationSubmission;
  if (!pending || !pending.id) {
    return;
  }
  console.log(`Deleting pending submission ${pending.id}...`);
  await deleteAppSubmission(`applications/${appId}/submissions/${pending.id}`);
}

/**
 * Sets the release notes of every listing in the submission.
 */
function setReleaseNotes(submissionResource: any, releaseNotes: string): void {
  var listings = submissionResource.listings || {};
  var keys = Object.keys(listings);
  if (keys.length === 0) {
    throw new Error("The submission has no listings to attach release notes to");
  }
  keys.forEach((listingKey) => {
    var listing = listings[listingKey];
    listing.baseListing = listing.baseListing || {};
    listing.baseListing.releaseNotes = releaseNotes;
    console.log(`Release notes set for listing ${listingKey}`);
  });
}

/**
 * Creates a submission for a given app.
 * @return Promises the new submission resource.
 */
function createAppSubmission(): Q.Promise<any> {
  return api.createSubmission(
    currentToken,
    api.ROOT + "applications/" + appId + "/submissions"
  );
}

/**
 * @return Promises the deletion of a resource
 */
function deleteAppSubmission(submissionLocation: string): Q.Promise<void> {
  return api.deleteSubmission(currentToken, api.ROOT + submissionLocation);
}

/**
 * @return Promises the resource associated with the application given to the task.
 */
async function getAppResource() {
  return api.getAppResource(currentToken, core.getInput("app-id"));
}

/**
 * Add any PendingUpload images in the given submission resource to the given zip file.
 */
function addImagesToZip(
  submissionResource: any,
  zip: { file: (arg0: any, arg1: any, arg2: { compression: string }) => void }
) {
  for (var listingKey in submissionResource.listings) {
    console.log(`Checking for new images in listing ${listingKey}...`);
    var listing = submissionResource.listings[listingKey];

    if (listing.baseListing.images) {
      addImagesToZipFromListing(listing.baseListing.images, zip);
    }

    for (var platOverrideKey in listing.platformOverrides) {
      console.log(
        `Checking for new images in platform override ${listingKey}/${platOverrideKey}...`
      );
      var platOverride = listing.platformOverrides[platOverrideKey];

      if (platOverride.images) {
        addImagesToZipFromListing(platOverride.images, zip);
      }
    }
  }
}

function addImagesToZipFromListing(
  images: any[],
  zip: { file: (arg0: any, arg1: any, arg2: { compression: string }) => void }
) {
  images
    .filter((image) => image.fileStatus == "PendingUpload")
    .forEach((image) => {
      var imgPath = path.join("", image.fileName);
      // According to JSZip documentation, the directory separator used is a forward slash.
      var filenameInZip = image.fileName.replace(/\\/g, "/");
      console.log(`Adding image path ${imgPath} to zip as ${filenameInZip}`);
      zip.file(filenameInZip, fs.createReadStream(imgPath), {
        compression: "DEFLATE",
      });
    });
}

/**
 * Commits a submission, checking for any errors.
 * @return A promise for the commit of the submission
 */
function commitAppSubmission(submissionId: string): Q.Promise<void> {
  return api.commitSubmission(
    currentToken,
    api.ROOT +
      "applications/" +
      appId +
      "/submissions/" +
      submissionId +
      "/commit"
  );
}

/**
 * Adds the required metadata to the submission request, depending on the given parameters.
 * If no metadata update is to be perfomed, no changes are made. Otherwise, we look for the metadata
 * depending on the type of update (text or json).
 * @param submissionResource The current submission request
 * @returns A promise for the update of the submission on the server.
 */
function putMetadata(submissionResource: any): Q.Promise<void> {
  console.log(`Adding metadata for new submission ${submissionResource.id}`);

  // Also at this point add the given packages to the list of packages to upload.
  api.includePackagesInSubmission(
    packages,
    submissionResource.applicationPackages
  );

  return putSubmission(submissionResource);
}

function putSubmission(submissionResource: any): Q.Promise<void> {
  var url =
    api.ROOT +
    "applications/" +
    appId +
    "/submissions/" +
    submissionResource.id;

  return api.putSubmission(currentToken, url, submissionResource);
}

async function main() {
  try {
    await publishTask();
  } catch (error) {
    core.setFailed(error.message);
  }
}

main();
