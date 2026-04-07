import test from "node:test";
import assert from "node:assert/strict";
import { copyFile, writeFile } from "node:fs/promises";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  addVideo,
  addAudio,
  getMediaElements,
  removeMediaElement,
  setMediaOptions,
} from "../src/media-advanced.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-media-adv-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

// Create a simple test video file (just bytes, not a real video)
function createTestVideoFile(ext: string = ".mp4"): Buffer {
  // Minimal valid MP4 file header (or just fake bytes for testing)
  // This is not a valid video file, but enough for testing embedding
  const header = Buffer.from([
    0x00, 0x00, 0x00, 0x20, 0x66, 0x74, 0x79, 0x70, // ftyp box
    0x69, 0x73, 0x6f, 0x6d, 0x00, 0x00, 0x02, 0x00, // isom
    0x69, 0x73, 0x6f, 0x6d, 0x69, 0x73, 0x6f, 0x32, // isom
    0x6d, 0x70, 0x34, 0x31, 0x00, 0x00, 0x00, 0x08, // mp41
    0x66, 0x72, 0x65, 0x65, // free
  ]);
  return Buffer.concat([header, Buffer.alloc(100, 0)]);
}

// Create a simple test audio file
function createTestAudioFile(ext: string = ".mp3"): Buffer {
  // Minimal MP3 frame header
  const frameHeader = Buffer.from([
    0xFF, 0xFB, 0x90, 0x00, // MP3 frame header
  ]);
  return Buffer.concat([frameHeader, Buffer.alloc(100, 0)]);
}

// Create a simple PNG image for poster
function createTestImage(): Buffer {
  const pngData = Buffer.from([
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54,
    0x08, 0xD7, 0x63, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x05, 0x00, 0x01,
    0x87, 0xA1, 0x4E, 0xD4,
    0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82
  ]);
  return pngData;
}

test("addVideo - adds a video to a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const videoData = createTestVideoFile();
    const videoPath = path.join(tmpdir(), `test-video-${Date.now()}.mp4`);
    await writeFile(videoPath, videoData);

    const result = await addVideo(
      tempPath,
      1,
      videoPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 2000000 },
      { autoplay: false, loop: false, mute: true }
    );

    assert.ok(result.ok, `addVideo failed: ${result.error?.message}`);
    assert.ok(result.data?.path);
    assert.ok(result.data?.path.includes("/slide[1]/media["));

    // Clean up temp video file
  } finally {
    // Clean up
  }
});

test("addVideo - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const videoData = createTestVideoFile();
    const videoPath = path.join(tmpdir(), `test-video-${Date.now()}.mp4`);
    await writeFile(videoPath, videoData);

    const result = await addVideo(
      tempPath,
      999,
      videoPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 2000000 }
    );

    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("addVideo - returns error for non-existent video file", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addVideo(
      tempPath,
      1,
      "/non/existent/video.mp4",
      { x: 1000000, y: 1000000, width: 3000000, height: 2000000 }
    );

    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("addVideo - adds video with poster image", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const videoData = createTestVideoFile();
    const videoPath = path.join(tmpdir(), `test-video-${Date.now()}.mp4`);
    await writeFile(videoPath, videoData);

    const posterData = createTestImage();
    const posterPath = path.join(tmpdir(), `test-poster-${Date.now()}.png`);
    await writeFile(posterPath, posterData);

    const result = await addVideo(
      tempPath,
      1,
      videoPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 2000000 },
      { autoplay: true, loop: true, posterImage: posterPath }
    );

    assert.ok(result.ok, `addVideo with poster failed: ${result.error?.message}`);
    assert.ok(result.data?.path);
  } finally {
    // Clean up
  }
});

test("addAudio - adds an audio file to a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const audioData = createTestAudioFile();
    const audioPath = path.join(tmpdir(), `test-audio-${Date.now()}.mp3`);
    await writeFile(audioPath, audioData);

    const result = await addAudio(
      tempPath,
      1,
      audioPath,
      { x: 1000000, y: 1000000, width: 500000, height: 500000 },
      { autoplay: true, loop: false, volume: 75 }
    );

    assert.ok(result.ok, `addAudio failed: ${result.error?.message}`);
    assert.ok(result.data?.path);
    assert.ok(result.data?.path.includes("/slide[1]/media["));
  } finally {
    // Clean up
  }
});

test("addAudio - adds audio without position (uses defaults)", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const audioData = createTestAudioFile();
    const audioPath = path.join(tmpdir(), `test-audio-${Date.now()}.mp3`);
    await writeFile(audioPath, audioData);

    const result = await addAudio(tempPath, 1, audioPath, undefined, { autoplay: false });

    assert.ok(result.ok, `addAudio without position failed: ${result.error?.message}`);
    assert.ok(result.data?.path);
  } finally {
    // Clean up
  }
});

test("addAudio - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const audioData = createTestAudioFile();
    const audioPath = path.join(tmpdir(), `test-audio-${Date.now()}.mp3`);
    await writeFile(audioPath, audioData);

    const result = await addAudio(
      tempPath,
      999,
      audioPath
    );

    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("getMediaElements - returns empty array for slide without video/audio", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getMediaElements(tempPath, 1);

    assert.ok(result.ok, `getMediaElements failed: ${result.error?.message}`);
    assert.ok(Array.isArray(result.data?.media));
    assert.equal(result.data?.total, 0);
  } finally {
    // Clean up
  }
});

test("getMediaElements - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getMediaElements(tempPath, 999);

    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("getMediaElements - returns video and audio on slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First add a video
    const videoData = createTestVideoFile();
    const videoPath = path.join(tmpdir(), `test-video-${Date.now()}.mp4`);
    await writeFile(videoPath, videoData);

    const videoResult = await addVideo(
      tempPath,
      1,
      videoPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 2000000 },
      { autoplay: true, loop: false }
    );
    assert.ok(videoResult.ok, `addVideo failed: ${videoResult.error?.message}`);

    // Then add an audio
    const audioData = createTestAudioFile();
    const audioPath = path.join(tmpdir(), `test-audio-${Date.now()}.mp3`);
    await writeFile(audioPath, audioData);

    const audioResult = await addAudio(
      tempPath,
      1,
      audioPath,
      { x: 2000000, y: 2000000, width: 500000, height: 500000 },
      { autoplay: false, volume: 50 }
    );
    assert.ok(audioResult.ok, `addAudio failed: ${audioResult.error?.message}`);

    // Get media elements
    const result = await getMediaElements(tempPath, 1);
    assert.ok(result.ok, `getMediaElements failed: ${result.error?.message}`);
    assert.ok(result.data);
    assert.ok(result.data?.media.length >= 2);

    // Check that we have at least one video and one audio
    const mediaTypes = result.data?.media.map(m => m.type) || [];
    assert.ok(mediaTypes.includes("video"), "Should have at least one video");
    assert.ok(mediaTypes.includes("audio"), "Should have at least one audio");
  } finally {
    // Clean up
  }
});

test("removeMediaElement - removes a video from a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First add a video
    const videoData = createTestVideoFile();
    const videoPath = path.join(tmpdir(), `test-video-${Date.now()}.mp4`);
    await writeFile(videoPath, videoData);

    const addResult = await addVideo(
      tempPath,
      1,
      videoPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 2000000 }
    );
    assert.ok(addResult.ok, `addVideo failed: ${addResult.error?.message}`);

    const videoPathResult = addResult.data!.path;

    // Now remove it
    const removeResult = await removeMediaElement(tempPath, videoPathResult);
    assert.ok(removeResult.ok, `removeMediaElement failed: ${removeResult.error?.message}`);
  } finally {
    // Clean up
  }
});

test("removeMediaElement - removes an audio from a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First add an audio
    const audioData = createTestAudioFile();
    const audioPath = path.join(tmpdir(), `test-audio-${Date.now()}.mp3`);
    await writeFile(audioPath, audioData);

    const addResult = await addAudio(
      tempPath,
      1,
      audioPath
    );
    assert.ok(addResult.ok, `addAudio failed: ${addResult.error?.message}`);

    const audioPathResult = addResult.data!.path;

    // Now remove it
    const removeResult = await removeMediaElement(tempPath, audioPathResult);
    assert.ok(removeResult.ok, `removeMediaElement failed: ${removeResult.error?.message}`);
  } finally {
    // Clean up
  }
});

test("removeMediaElement - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeMediaElement(tempPath, "/slide[1]/media[999]");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("removeMediaElement - returns error for path without media index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeMediaElement(tempPath, "/slide[1]");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("setMediaOptions - updates video options", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First add a video
    const videoData = createTestVideoFile();
    const videoPath = path.join(tmpdir(), `test-video-${Date.now()}.mp4`);
    await writeFile(videoPath, videoData);

    const addResult = await addVideo(
      tempPath,
      1,
      videoPath,
      { x: 1000000, y: 1000000, width: 3000000, height: 2000000 },
      { autoplay: false, loop: false, mute: true }
    );
    assert.ok(addResult.ok, `addVideo failed: ${addResult.error?.message}`);

    const videoPathResult = addResult.data!.path;

    // Update options
    const updateResult = await setMediaOptions(tempPath, videoPathResult, {
      autoplay: true,
      loop: true,
      mute: false,
      volume: 80,
    });

    assert.ok(updateResult.ok, `setMediaOptions failed: ${updateResult.error?.message}`);

    // Verify the update
    const getResult = await getMediaElements(tempPath, 1);
    assert.ok(getResult.ok, `getMediaElements failed: ${getResult.error?.message}`);

    const updatedMedia = getResult.data?.media.find(m => m.path === videoPathResult);
    assert.ok(updatedMedia);
    assert.equal(updatedMedia?.autoplay, true);
    assert.equal(updatedMedia?.loop, true);
    assert.equal(updatedMedia?.mute, false);
    assert.equal(updatedMedia?.volume, 80);
  } finally {
    // Clean up
  }
});

test("setMediaOptions - updates audio options", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First add an audio
    const audioData = createTestAudioFile();
    const audioPath = path.join(tmpdir(), `test-audio-${Date.now()}.mp3`);
    await writeFile(audioPath, audioData);

    const addResult = await addAudio(
      tempPath,
      1,
      audioPath,
      undefined,
      { autoplay: false, loop: false, volume: 50 }
    );
    assert.ok(addResult.ok, `addAudio failed: ${addResult.error?.message}`);

    const audioPathResult = addResult.data!.path;

    // Update options
    const updateResult = await setMediaOptions(tempPath, audioPathResult, {
      autoplay: true,
      loop: true,
      volume: 90,
    });

    assert.ok(updateResult.ok, `setMediaOptions failed: ${updateResult.error?.message}`);

    // Verify the update
    const getResult = await getMediaElements(tempPath, 1);
    assert.ok(getResult.ok, `getMediaElements failed: ${getResult.error?.message}`);

    const updatedMedia = getResult.data?.media.find(m => m.path === audioPathResult);
    assert.ok(updatedMedia);
    assert.equal(updatedMedia?.autoplay, true);
    assert.equal(updatedMedia?.loop, true);
    assert.equal(updatedMedia?.volume, 90);
  } finally {
    // Clean up
  }
});

test("setMediaOptions - returns error for non-existent media", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setMediaOptions(
      tempPath,
      "/slide[1]/media[999]",
      { autoplay: true }
    );

    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("setMediaOptions - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setMediaOptions(
      tempPath,
      "/slide[1]",
      { autoplay: true }
    );

    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});
