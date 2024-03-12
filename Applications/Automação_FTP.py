import paramiko
import sys

if getattr(sys, 'frozen', False):
    sys.argv = [sys.executable] + sys.argv

host = ''
port = 22
username = ''
password = ''

local_directory = ''

remote_directory = '/root/dados/relatorio'

def download_latest_file():
    try:
        transport = paramiko.Transport((host, port))
        transport.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)

        files = [file for file in sftp.listdir(remote_directory) if "IAR" in file]

        latest_file = None
        latest_timestamp = None

        for file in files:
            timestamp_str = file.split('_')[2].split('.')[0]
            timestamp = int(timestamp_str)

            if latest_timestamp is None or timestamp > latest_timestamp:
                latest_file = file
                latest_timestamp = timestamp

        if latest_file:
            remote_file_path = remote_directory + '/' + latest_file
            local_file_path = local_directory + '/' + latest_file
            sftp.get(remote_file_path, local_file_path)

        sftp.close()
        transport.close()

    except Exception as e:
        print(f"Erro ao baixar arquivos: {e}")

if __name__ == '__main__':
    download_latest_file()