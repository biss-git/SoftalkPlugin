using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yomiage.SDK;
using Yomiage.SDK.Config;
using Yomiage.SDK.Talk;
using Yomiage.SDK.VoiceEffects;

namespace SoftalkPlugin
{
    public class VoiceEngine : VoiceEngineBase
    {
        private Process process;

        private string textFile => Path.Combine(DllDirectory, "text.txt");
        private string wavFile => Path.Combine(DllDirectory, "output.wav");


        public override void Initialize(string configDirectory, string dllDirectory, EngineConfig config)
        {
            base.Initialize(configDirectory, dllDirectory, config);
            StartUp();
        }

        public override async Task<double[]> Play(VoiceConfig mainVoice, VoiceConfig subVoice, TalkScript talkScript, MasterEffectValue masterEffect, Action<int> setSamplingRate_Hz, Action<double[]> submitWavePart)
        {
            if (!StartUp())
            {
                return new double[0];
            }

            if (File.Exists(textFile)) { File.Delete(textFile); }
            if (File.Exists(wavFile)) { File.Delete(wavFile); }

            var speed = (int)Math.Round(mainVoice.VoiceEffect.Speed.GetValueOrDefault(100) * masterEffect.Speed.GetValueOrDefault(1) * talkScript.Speed.GetValueOrDefault(1));

            SetText(
                talkScript,
                textFile,
                Settings.GetAccentEnable(),
                speed: speed);

            SaveVoice(
                exePath: Settings.GetExePath(),
                textPath: textFile,
                outputPath: wavFile,
                volume: (int)Math.Round(mainVoice.VoiceEffect.Volume.GetValueOrDefault(50) * masterEffect.Volume.GetValueOrDefault(1) * talkScript.Volume.GetValueOrDefault(1)),
                speed: speed,
                pitch: (int)Math.Round(mainVoice.VoiceEffect.Pitch.GetValueOrDefault(50) * masterEffect.Pitch.GetValueOrDefault(1) * talkScript.Pitch.GetValueOrDefault(1)),
                accent: (int)Math.Round(mainVoice.VoiceEffect.Emphasis.GetValueOrDefault(100) * masterEffect.Emphasis.GetValueOrDefault(1) * talkScript.Emphasis.GetValueOrDefault(1)),
                interval: (int)Math.Round(mainVoice.VoiceEffect.AdditionalEffect.GetValueOrDefault("Interval").GetValueOrDefault(100) * talkScript.AdditionalEffect.GetValueOrDefault("Interval").GetValueOrDefault(1)),
                interval2: (int)Math.Round(mainVoice.VoiceEffect.AdditionalEffect.GetValueOrDefault("Interval2").GetValueOrDefault(100)),
                voiceName: mainVoice.Library.Settings.GetVoiceName(),
                presetName: mainVoice.Library.Settings.GetPresetName()
                );

            return await GetVoice(
                wavFile,
                setSamplingRate_Hz,
                talkScript.Sections.First().Pause.Span_ms,
                talkScript.EndSection.Pause.Span_ms + (int)masterEffect.EndPause
                );
        }

        public override void Dispose()
        {
            base.Dispose();
            process?.Kill();
            process?.Dispose();
        }

        /// <summary>
        /// Softalkが起動しているか確認する。
        /// また、ExePathが設定されているか確認する
        /// 起動していなければ起動する。
        /// </summary>
        private bool StartUp()
        {
            var path = Settings?.GetExePath();
            if (string.IsNullOrWhiteSpace(path) ||
                !File.Exists(path))
            {
                // Softalk のパスが設定されていない
                this.StateText = "Saftalk のパスが設定されていません。";

                var processList = Process.GetProcessesByName("SofTalk");
                if (processList.Count() == 0)
                {
                    return false;
                }

                var fileName = processList.First().MainModule.FileName;
                if (!File.Exists(fileName) ||
                    !fileName.Contains("softalk", StringComparison.OrdinalIgnoreCase) ||
                    Settings?.SetExePath(fileName) != true)
                {
                    return false;
                }
                path = fileName;
            }

            if (process?.HasExited == false)
            {
                // 既にプロセスを自分で起動している
                return true;
            }

            if(Process.GetProcessesByName("SofTalk").Count() > 0)
            {
                // 既にプロセスが起動していた
                return true;
            }

            var processStartInfo = new ProcessStartInfo()
            {
                FileName = path,
                ArgumentList =
                {
                    $"/X:{1}",  // 画面の非表示化
                },
            };
            process = Process.Start(processStartInfo);
            return true;
        }

        private void SetText(TalkScript talkScript, string textPath, bool accentEnable, int speed)
        {
            var text = talkScript.OriginalText;
            if (accentEnable)
            {
                text = talkScript.GetYomiForAquesTalkLike();
                speed = Math.Clamp(speed, 10, 300);

                var list = text.Split('/');
                text = list.First();
                for (int i = 1; i < Math.Min(talkScript.Sections.Count, list.Length); i++)
                {
                    var section = talkScript.Sections[i];
                    if(section.Pause.Span_ms > 0)
                    {
                        var pauseNumber = section.Pause.Span_ms * speed / 10000;
                        pauseNumber = Math.Clamp(pauseNumber, 1, 30);
                        text += new string(',', pauseNumber);
                    }
                    else
                    {
                        text += "/";
                    }
                    text += list[i];
                }
            }
            File.WriteAllText(textPath, text);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="exePath">Softalkのexeパス</param>
        /// <param name="textPath">読み上げる対象のテキスト</param>
        /// <param name="outputPath">出力先パス</param>
        /// <param name="volume">音量 0-100</param>
        /// <param name="speed">速度 1-300</param>
        /// <param name="pitch">高さ 20-200 AT10</param>
        /// <param name="accent">ｱｸｾﾝﾄ 0-200 AT10</param>
        /// <param name="interval">音程 0-300</param>
        /// <param name="interval2">音程2 50-200 AT10</param>
        /// <param name="voiceName">話者名</param>
        /// <param name="presetName">プリセット名</param>
        private void SaveVoice(
            string exePath,
            string textPath,
            string outputPath,
            int volume,
            int speed,
            int pitch,
            int accent,
            int interval,
            int interval2,
            string voiceName,
            string presetName
            )
        {
            volume = Math.Clamp(volume, 0, 100);
            speed = Math.Clamp(speed, 1, 300);
            pitch = Math.Clamp(pitch, 20, 200);
            accent = Math.Clamp(accent, 0, 200);
            interval = Math.Clamp(interval, 0, 300);
            interval2 = Math.Clamp(interval2, 50, 200);

            var processStartInfo = new ProcessStartInfo()
            {
                FileName = exePath,
                ArgumentList =
                {
                    textPath, // 読み取りパス
                    $"/R:{outputPath}", // 保存パス
                    $"/PS:True", // 再生しないかどうか
                },
                CreateNoWindow = true,
                UseShellExecute = false,
            };
            if (!string.IsNullOrWhiteSpace(voiceName) && voiceName != "None")
            {
                processStartInfo.ArgumentList.Add($"/NM:{voiceName}");
            }
            if (!string.IsNullOrWhiteSpace(presetName))
            {
                processStartInfo.ArgumentList.Add($"/PR:{presetName}");
            }

            processStartInfo.ArgumentList.Add($"/J:{pitch}"); // 高さ 20-200 AT10
            processStartInfo.ArgumentList.Add($"/K:{accent}"); // ｱｸｾﾝﾄ 0-200 AT10
            processStartInfo.ArgumentList.Add($"/L:{interval2}"); // 音程2 50-200 AT10
            processStartInfo.ArgumentList.Add($"/O:{interval}"); // 音程 0-300
            processStartInfo.ArgumentList.Add($"/S:{speed}"); // 速度 1-300
            processStartInfo.ArgumentList.Add($"/V:{volume}"); // 音量 0-100

            Process.Start(processStartInfo);
        }


        private async Task<double[]> GetVoice(
            string wavPath,
            Action<int> setSamplingRate_Hz,
            int startPause_ms,
            int endPause_ms
            )
        {
            var fs = 44100;
            var wave = new List<double>();
            for(int i = 0; i < 10; i++)
            {
                await Task.Delay(100);
                if (File.Exists(wavPath))
                {
                    using var reader = new WaveFileReader(wavPath);
                    fs = reader.WaveFormat.SampleRate;
                    setSamplingRate_Hz(fs);
                    while (reader.Position < reader.Length)
                    {
                        var samples = reader.ReadNextSampleFrame();
                        wave.Add(samples.First());
                    }
                    break;
                }
            }

            if (wave.Count == 0)
            {
                return new double[0];
            }

            wave.InsertRange(0, new double[fs * startPause_ms / 1000]);
            wave.AddRange(new double[fs * endPause_ms / 1000]);

            return wave.ToArray();
        }
    }
}
