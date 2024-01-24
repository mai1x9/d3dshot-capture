import threading
import collections
from PIL import ImageGrab, Image
from io import BytesIO
from time import time, sleep


try:
    from .src import dxgi
except ImportError:
    pass


__all__ = ["ScreenRecordDupAPI", "DirectScreenRecord"]  # Import functions


def raw_to_memory(raw_bytes, width, height, region=None, hd="1080p", quality=75, memory=False):
    # hd = full means 1920*1080p which is 1080p
    image = Image.frombytes("RGBA", (width, height), raw_bytes)
    b, g, r, _ = image.split()
    image = Image.merge("RGB", (r, g, b))

    # Trim pitch padding
    # image = image.crop((0, 0, width, height))

    # Region slicing
    if region:
        if region[2] - region[0] != width or region[3] - region[1] != height:
            image = image.crop(region)

    if hd == "720p":
        # 720p
        # Do not lower quality less than 40% for 720p, can go upto 30% for avg quality
        image = image.resize((1280, 720), Image.ANTIALIAS)  # HD

    if hd == "480p":
        # 480p, do not lower than 80%
        image = image.resize((854, 480), Image.ANTIALIAS)  # medium

    if hd == "360p":
        # 360p, set max quality
        image = image.resize((640, 360), Image.ANTIALIAS)  # medium

    if hd == "240p":
        # 360p, set max quality
        image = image.resize((426, 240), Image.ANTIALIAS)  # medium

    if hd == "144p":
        # 360p, set max quality
        image = image.resize((256, 144), Image.ANTIALIAS)  # medium

    if memory:
        # save in memory
        mem = BytesIO()
        image.save(mem, "jpeg", quality=quality)
        # For full HD, quality can be set till 20%. In worst case it can go till 10% and still looks good.
        # for 720p, do not less quality than 25%
        return mem.getvalue()

    # return PIL image
    return image


def pil_to_memory(image, width, height, region=None, hd="1080p", quality=75):
    # Region slicing
    if region:
        if region[2] - region[0] != width or region[3] - region[1] != height:
            image = image.crop(region)

    if hd == "720p":
        # 720p
        # Do not lower quality less than 40% for 720p, can go upto 30% for avg quality
        image = image.resize((1280, 720), Image.ANTIALIAS)  # HD

    if hd == "480p":
        # 480p, do not lower than 80%
        image = image.resize((854, 480), Image.ANTIALIAS)  # medium

    if hd == "360p":
        # 360p, set max quality
        image = image.resize((640, 360), Image.ANTIALIAS)  # medium

    if hd == "240p":
        # 360p, set max quality
        image = image.resize((426, 240), Image.ANTIALIAS)  # medium

    if hd == "144p":
        # 360p, set max quality
        image = image.resize((256, 144), Image.ANTIALIAS)  # medium

    # save in memory
    mem = BytesIO()
    image.save(mem, "jpeg", quality=quality)
    # For full HD, quality can be set till 20%. In worst case it can go till 10% and still looks good.
    # for 720p, do not less quality than 25%
    return mem.getvalue()


class Display:
    def __init__(self):
        self.primary = None

        self.dxgi_output_duplication = None
        self.d3d_device = None
        self.width = None
        self.height = None

        display_device_name_mapping = dxgi.get_display_device_name_mapping()
        dxgi_factory = dxgi.initialize_dxgi_factory()
        dxgi_adapters = dxgi.discover_dxgi_adapters(dxgi_factory)

        for dxgi_adapter in dxgi_adapters:
            for dxgi_output in dxgi.discover_dxgi_outputs(dxgi_adapter):
                dxgi_output_description = dxgi.describe_dxgi_output(dxgi_output)

                if dxgi_output_description["is_attached_to_desktop"]:
                    display_device = display_device_name_mapping.get(dxgi_output_description["name"])

                    if display_device is None:
                        continue

                    self.primary = display_device[1]  # Only Primary window is accepted

                    # Set resolutions
                    resolution = dxgi_output_description["resolution"]
                    self.width = resolution[0]
                    self.height = resolution[1]

                    self.d3d_device = dxgi.initialize_d3d_device(dxgi_adapter)[0]
                    self.dxgi_output_duplication = dxgi.initialize_dxgi_output_duplication(
                        dxgi_output, self.d3d_device)

        print(self.primary, self.width, self.height, self.d3d_device)

    def desktop_dup_api(self, resolution=None):
        frame = None

        if not self.primary:
            return None

        if resolution is None:
            resolution = (self.width, self.height)

        try:
            frame = dxgi.get_dxgi_output_duplication_frame(
                self.dxgi_output_duplication, self.d3d_device, height=resolution[1])
        except Exception:
            pass

        return frame


class ScreenRecordDupAPI:
    # Desktop duplication API
    def __init__(self, frame_buffer_size=180, region=None, memory=True):
        # 180 frame roughly 10 second vedio with avg size of 8 Mb in memory
        self.display = Display()
        self.width = self.display.width
        self.height = self.display.height
        self.frame_buffer_size = frame_buffer_size
        self.frame_buffer = collections.deque(list(), self.frame_buffer_size)
        # Keep appending to deque, consumer will pop from left from the queue
        # Producer will keep appending vedio frames
        # When maximum length reached, the first(old vedio frames) will be automatically
        # be removed from the queue
        self.fps = 15
        self.region = region  # region to be captured
        self.memory = memory  # if True, store frames in memory
        self._is_capturing = False

    def get_frame_buffer(self):
        # pop each frame, the first frame into the queue, will be the first one to be processed
        # Eg: sending over network etc ...
        return self.frame_buffer.popleft()

    def screenshot(self):
        frame = None
        while frame is None:
            frame = self.display.desktop_dup_api()

        frame = raw_to_memory(frame, self.width, self.height, memory=True, quality=30, hd="720p")
        return frame

    def capture(self, fps=15, hd="1080p", quality=75):
        # runs on seperate thread, at any time only once you can launch capture
        if self._is_capturing:
            return False

        self.fps = fps
        self._is_capturing = True
        threading.Thread(target=self._capture, args=(hd, quality,)).start()
        return True

    def _capture(self, hd, quality):
        frame_time = 1 / self.fps
        region = self.region
        memory = self.memory

        while self._is_capturing:
            start = time()

            frame = self.display.desktop_dup_api()

            if frame is not None:
                frame = raw_to_memory(frame, self.width, self.height,
                                      region=region, hd=hd, quality=quality, memory=memory)
                print("Frame details ....")
                self.frame_buffer.appendleft(frame)
            else:
                if len(self.frame_buffer):
                    self.frame_buffer.appendleft(self.frame_buffer[-1])

            now = time()
            # gc.collect()  # can be removed, check if any performance impact or not

            frame_time_left = frame_time - (now - start)
            if frame_time_left > 0:
                sleep(frame_time_left)

        self._is_capturing = False

    def stop(self):
        # call stop(), to stop previous recording. Now you can again start a new recording by calling
        # capture()
        if not self._is_capturing:
            return False

        self._is_capturing = False
        return True

    # def capture_a(self, target_fps=15):
    #     # self.frame_buffer = collections.deque(list(), self.frame_buffer_size)
    #     frame_time = 1 / target_fps
    #     print("fps => ", target_fps)
    #
    #     while 1:
    #         self.frame_buffer = collections.deque(list(), self.frame_buffer_size)
    #         a = time()
    #         count = 0
    #         print("frame buff init ", len(self.frame_buffer))
    #
    #         while time() - a <= 1:
    #             cycle_start = time()
    #             frame = self.display.desktop_dup_api()
    #             if frame is not None:
    #                 # frame = to_image(frame, self.width, self.height, memory=True)
    #                 self.frame_buffer.append(frame)
    #             else:
    #                 if len(self.frame_buffer):
    #                     self.frame_buffer.append(self.frame_buffer[0])
    #
    #             cycle_end = time()
    #             # gc.collect()  # can be removed, check if any performance impact or not
    #             frame_time_left = frame_time - (cycle_end - cycle_start)
    #             if frame:
    #                 count += 1
    #                 print("frame timeleft ", frame_time_left, " ...count ", count, " ... frame szie ", len(frame))
    #             # else:
    #             #     print("frame timeleft ", frame_time_left , " ...count ", count )
    #
    #             if frame_time_left > 0:
    #                sleep(frame_time_left)
    #
    #         print("End of loop ...size of frames ", len(self.frame_buffer), " count => ", count)
    #
    #     self._is_capturing = False


class DirectScreenRecord:
    # Direct X11 using ImageGrab
    def __init__(self, frame_buffer_size=180, region=None, memory=True):
        self.frame_buffer_size = frame_buffer_size
        self.framebuffer = collections.deque(list(), self.frame_buffer_size)

        self.width, self.height = ImageGrab.grab().size

        self.fps = 15
        self.region = region
        self.memory = memory  # store frames in memory
        self._is_capturing = False

    def get_frame_buffer(self):
        # pop each frame, the first frame into the queue, will be the first one to be processed
        # Eg: sending over network etc ...
        return self.frame_buffer.popleft()

    def capture(self, fps=15, hd="1080p", quality=75):
        if self._is_capturing:
            return False

        self.fps = fps
        self._is_capturing = True
        threading.Thread(target=self._capture, args=(hd, quality,)).start()
        return True

    def _capture(self, hd, quality):
        frame_time = 1 / self.fps
        region = self.region

        while self._is_capturing:
            start = time()
            image = ImageGrab.grab()
            frame = pil_to_memory(image, self.width, self.height, region=region, hd=hd, quality=quality)
            self.framebuffer.append(frame)
            print("-----------")
            now = time()

            frame_time_left = frame_time - (now - start)

            if frame_time_left > 0:
                sleep(frame_time_left)

        self._is_capturing = False
        return True

    def stop(self):
        if not self._is_capturing:
            return False

        self._is_capturing = False
        return True


"""
API Use:
rec = record.ScreenRecordDupAPI()
rec.capture()


rec = record.DirectScreenRecord()
rec.capture()

"""