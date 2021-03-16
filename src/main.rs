fn main() -> anyhow::Result<()> {
    #[cfg(not(windows))]
    unimplemented!();

    use volume::{Channel, Volume};

    let args: Vec<String> = std::env::args().collect();
    let mut vol_left = 100usize;
    let mut vol_right = 86usize;
    if args.len() >= 3 {
        vol_left = args[1].parse::<usize>()?;
        vol_right = args[2].parse::<usize>()?;
    }
    let mut volume = Volume::new();
    volume.initialize()?;
    let left_volume = volume.get_channel_volume(Channel::Left as u32)?;
    let right_volume = volume.get_channel_volume(Channel::Right as u32)?;
    let master_volume = volume.get_master_volume()?;
    if left_volume == right_volume {
        volume.set_channel_volume(Channel::Left as u32, vol_left)?;
        volume.set_channel_volume(Channel::Right as u32, vol_right)?;
    } else {
        volume.set_channel_volume(Channel::Left as u32, 100)?;
        volume.set_channel_volume(Channel::Right as u32, 100)?;
    }
    volume.set_master_volume(master_volume)?;

    Ok(())
}

#[cfg(windows)]
mod volume {
    use anyhow::{anyhow, Result};
    use std::ptr::null_mut;
    use winapi::{
        ctypes::c_void,
        shared::{winerror::IS_ERROR, wtypesbase::CLSCTX_INPROC_SERVER},
        um::{
            combaseapi::{CoCreateInstance, CoInitializeEx, CLSCTX_ALL},
            endpointvolume::IAudioEndpointVolume,
            mmdeviceapi::{IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator},
            objbase::COINIT_MULTITHREADED,
            winnt::HRESULT,
        },
        Class, Interface,
    };

    pub enum Channel {
        Left = 0,
        Right = 1,
    }

    pub struct Volume<'a> {
        endpoint_volume: Option<&'a IAudioEndpointVolume>,
    }

    impl<'a> Volume<'a> {
        pub fn new() -> Self {
            Self {
                endpoint_volume: None,
            }
        }

        pub fn initialize(&mut self) -> Result<()> {
            let mut device: *mut IMMDevice = null_mut();
            let mut device_enumerator: *mut IMMDeviceEnumerator = null_mut();
            let mut endpoint_volume: *mut IAudioEndpointVolume = null_mut();
            let device_cast: *mut *mut IMMDevice = &mut device as *mut *mut _;
            let device_enumerator_void_cast: *mut *mut c_void =
                &mut device_enumerator as *mut _ as *mut *mut c_void;
            let endpoint_volume_void_cast: *mut *mut c_void =
                &mut endpoint_volume as *mut _ as *mut *mut c_void;

            let mut hr: HRESULT;
            hr = unsafe { CoInitializeEx(null_mut(), COINIT_MULTITHREADED) };
            is_error(hr, "CoInitializeEx")?;
            hr = unsafe {
                CoCreateInstance(
                    &MMDeviceEnumerator::uuidof(),
                    null_mut(),
                    CLSCTX_INPROC_SERVER,
                    &IMMDeviceEnumerator::uuidof(),
                    device_enumerator_void_cast,
                )
            };
            is_error(hr, "CoCreateInstance")?;
            is_null(device_enumerator)?;
            unsafe {
                let device_enumerator = &*device_enumerator;
                device_enumerator.GetDefaultAudioEndpoint(0u32, 1u32, device_cast);
            };
            is_null(device)?;
            unsafe {
                let device = &*device;
                hr = device.Activate(
                    &IAudioEndpointVolume::uuidof(),
                    CLSCTX_ALL,
                    null_mut(),
                    endpoint_volume_void_cast,
                );
            };
            is_error(hr, "Activate")?;
            is_null(endpoint_volume)?;
            let endpoint_volume = unsafe { &*endpoint_volume };
            self.endpoint_volume = Some(endpoint_volume);
            Ok(())
        }

        pub fn get_master_volume(&self) -> Result<usize> {
            let endpoint_volume = self.get_endpoint_volume()?;
            let mut volume = 0f32;
            let hr: HRESULT = unsafe { endpoint_volume.GetMasterVolumeLevelScalar(&mut volume) };
            is_error(hr, "GetMasterVolumeLevelScalar")?;
            let volume = volume * 100f32;
            Ok(volume.round() as usize)
        }

        pub fn get_channel_volume(&self, channel: u32) -> Result<usize> {
            let endpoint_volume = self.get_endpoint_volume()?;
            let mut volume = 0f32;
            let hr: HRESULT =
                unsafe { endpoint_volume.GetChannelVolumeLevelScalar(channel, &mut volume) };
            is_error(hr, "GetChannelVolumeLevelScalar")?;
            let volume = volume * 100f32;
            Ok(volume.round() as usize)
        }

        pub fn set_master_volume(&self, volume: usize) -> Result<()> {
            let endpoint_volume = self.get_endpoint_volume()?;
            if volume > 100 {
                return Err(anyhow!("Volume can't be above 100!".to_string()));
            }
            let volume: f32 = volume as f32 / 100f32;
            let hr: HRESULT =
                unsafe { endpoint_volume.SetMasterVolumeLevelScalar(volume, null_mut()) };
            is_error(hr, "SetMasterVolumeLevelScalar")?;
            Ok(())
        }

        pub fn set_channel_volume(&self, channel: u32, volume: usize) -> Result<()> {
            let endpoint_volume = self.get_endpoint_volume()?;
            if volume > 100 {
                return Err(anyhow!("Volume can't be above 100!".to_string()));
            }
            let volume: f32 = volume as f32 / 100f32;
            let hr: HRESULT =
                unsafe { endpoint_volume.SetChannelVolumeLevelScalar(channel, volume, null_mut()) };
            is_error(hr, "SetChannelVolumeLevelScalar")?;
            Ok(())
        }

        fn get_endpoint_volume(&self) -> Result<&IAudioEndpointVolume> {
            if let Some(vol) = &self.endpoint_volume {
                Ok(vol)
            } else {
                Err(anyhow!("Call initialize function first!".to_string()))
            }
        }
    }

    fn is_error(hr: HRESULT, name: &str) -> Result<()> {
        if IS_ERROR(hr) {
            Err(anyhow!(format!("Func: {}, HRESULT: {:#x}", name, hr)))
        } else {
            Ok(())
        }
    }

    fn is_null<T>(ptr: *mut T) -> Result<()> {
        if ptr.is_null() {
            Err(anyhow!(format!(
                "Pointer of type {:?} is null!",
                std::any::type_name::<T>()
            )))
        } else {
            Ok(())
        }
    }
}
