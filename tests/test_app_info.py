import unittest
import wave

from app_info import NOTIFICATION_SOUND, play_notification_sound


class AppInfoTests(unittest.TestCase):
    def test_som_de_notificacao_existe_e_e_wav_valido(self):
        self.assertTrue(NOTIFICATION_SOUND.is_file())
        with wave.open(str(NOTIFICATION_SOUND), 'rb') as audio:
            self.assertGreater(audio.getnframes(), 0)
            self.assertGreater(audio.getframerate(), 0)

    def test_reproducao_do_som_nao_interrompe_o_programa(self):
        self.assertIsInstance(play_notification_sound(), bool)


if __name__ == '__main__':
    unittest.main()
