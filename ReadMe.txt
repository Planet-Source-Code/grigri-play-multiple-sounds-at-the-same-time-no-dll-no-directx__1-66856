Grigri's Sound Manager for VB6
==============================

Written by
  : grigri <grigri@shinyhappypixels.com>
  
Version
  : 1.0 [22/10/2006]
  
Intro
-----
  
Playing multiple sounds simultaneously is not an easy task, especially in VB.
Most VB applications (and games!) that use sound use the PlaySound()
API function [which only allows one sound at a time]; others resort to DirectX
or third-party DLLs such as BASS. This project is a first step to ending all
that, presenting an API-only method.

The multimedia API set is nasty, really nasty. Even in C it's not easy to
manage, and in VB it just gets worse. I have waded through most of the ambiguous
'documentation' and the pathetic 'code examples' and I believe I've made a
good start.

How It Works
------------

The system is deployed as one module and a notification interface.
An internal array of sound buffers (limited to 32 but easily changed) is managed
by the system. All external calls reference the buffer index.

Confused? Here's a really quick-and-dirty example of how to use it.

    ' Load, play and free a sound in one line of code
    LoadSoundFile FreeBuffer, "c:\some\sound\file.wav", BufferFlagInstant

Important Point: At the end of the application, you **MUST** call
`SoundManager.DestroySoundManager()`. If you don't, you'll end up crashing your
application or VB, depending on whether you're in the IDE or not.

The notification interface allows you to be notified when a sound is loaded,
freed, playing or stopped.
    
The demo application shows the proper usage.

SoundManager API
----------------

* `DestroySoundManager()`
  Parameters  : None
  Return Type : None (`Sub`)
  
  Frees all loaded sounds and destroys internal structures. This *MUST* be
  called when you're done. I can't stress this enough.
  
* `FreeBuffer()`
  Parameters  : None
  Return Type : `Long`
  
  Returns the first free buffer index. Analogous to the VB `FreeFile()` function
  
* `SoundStatus()`
  Parameters  : - `BufferIndex` (IN)          [`Long`]
  Return Type : `SoundBufferStatus` (Enumeration)
  
  Returns the status of the buffer `BufferIndex`

* `LoadSoundFile()`
  Parameters  : - `BufferIndex` (IN)          [`Long`]
                - `FileName`    (IN)          [`String`]
                - `Flags`       (IN,OPTIONAL) [`SoundBufferFlags` (Enumeration)]
  Return Type : `Boolean`

  Loads the specified wave file into the specified buffer. The optional flags
  can be used to make the buffer play/free itself automatically, and/or to not
  generate notifications of status change.
  If the specified buffer is not empty, it will be stopped/freed as required.
  
* `FreeSound()`
  Parameters  : - `BufferIndex` (IN)          [`Long`]
  Return Type : None (`Sub`)
  
  Frees the specified buffer. If it is currently playing, it will be stopped.

* `StopSound()`
  Parameters  : - `BufferIndex` (IN)          [`Long`]
  Return Type : None (`Sub`)

  Stops playback on the specified buffer. If not playing, nothing happens.
  
* `PlaySound()`
  Parameters  : - `BufferIndex` (IN)          [`Long`]
  Return Type : `Boolean`

  Begins playback of the specified buffer. If it is currently playing, it will
  be stopped first, resulting in a "restart" -- playing the buffer from the
  beginning.

Notification Interface Callback Methods
---------------------------------------

All callback methods have identical prototypes, passing the buffer index as the
sole parameter, and not returning any value.

Each method corresponds to a status change of the specified buffer, and is
called *after* the status has been updated.

The names are self-explanatory.

* `SoundLoaded()`    : Status is `BufferLoaded`
* `SoundUnloaded()`  : Status is `BufferEmpty`
* `SoundPlayStart()` : Status is `BufferPlaying`
* `SoundPlayEnd()`   : Status is `BufferLoaded`

Error Handling
--------------

Most functions will simply fail silently in case of an error, returning a
`False` value, if appropriate.

The current exception is the `LoadSoundFile()` function, which uses
`MsgBox()` to display errors.

Better error handling is being planned for the next version. Honestly.
  
Limitations
-----------

* This version only supports uncompressed .WAV files. No .mp3, .au, .voc, .ogg
  or .aac.
  
* Every sound to be played is loaded in memory entirely before playback begins.
  This means that large sound files are not catered for by this method.
  Streaming sounds IS possible with the API, using a free-threaded rotating
  buffer chain, which seems to be impossible (or at least EXTREMELY difficult)
  in VB.
  
* Although it seems stable now, it was quite unstable during development and I
  had more than a few crashes. Each crash took VB down with it.
  So be warned! Use at your own risk.

Future Enhancements Planned (short-term)
---------------------------

* Better error handling
* Pause/Resume playback
* Looping
* Loading sounds from resources and memory
* Volume and pitch control

Future Enhancements Hoped For (long-term)
-----------------------------

* Streaming playback
* Support for different sound formats
