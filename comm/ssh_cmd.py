import time
import logging
import paramiko
from modbus_config import modbus_config


def exec_cmd(cmd):
    """
    执行单条命令，只负责执行
    :param cmd: 执行命令
    :return:
    """
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(username=modbus_config['ssh']['username'], password=modbus_config['ssh']['password'],
                   hostname=modbus_config['ssh']['ip'], port=modbus_config['ssh']['port'])

    stdin, stdout, stderr = client.exec_command(cmd, get_pty=True)
    client.close()
    # return str(stdout.read().decode('utf-8'))
    return stdout.readline()


def send_cmd(cmd):
    """
    执行单挑命令，并返回执行结果
    :param cmd: 命令
    :return: 执行结果
    """
    cmd_exec = cmd + '\n'
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(username=modbus_config['ssh']['username'], password=modbus_config['ssh']['password'],
                   hostname=modbus_config['ssh']['ip'], port=modbus_config['ssh']['port'])
    command = client.invoke_shell()
    command.send(cmd_exec)
    time.sleep(1)
    output = command.recv(65535)
    client.close()
    logging.info('cmd is:{}, exec recard is:{}'.format(cmd, output))
    return output.decode('utf-8').split('~#')[1].split('root@S8P54040023')[0].split(cmd)[1]


def send_cmds(cmds):
    """
    执行多条命令，并返回执行结果
    :param cmds: 多条命令，eg:
            '''
            command 1
            command 2
            ...
            '''
    :return:执行结果，[command 1 result, command 2 result, ...]
    """
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(username=modbus_config['ssh']['username'], password=modbus_config['ssh']['password'],
                   hostname=modbus_config['ssh']['ip'], port=modbus_config['ssh']['port'])
    command = client.invoke_shell()
    outputs = []
    cmds = [element.strip() for element in cmds.split('\n') if element.strip() != '']
    for cmd in cmds:
        command.send(cmd + '\n')
        time.sleep(1)
        output = command.recv(65535)
        logging.info('cmd is:{}, exec recard is:{}'.format(cmd, output))
        if len(outputs) == 0:
            output = output.decode('utf-8').split('~#')[1]
            output = output.split('root@S8P54040023')[0]
            output = output.split(cmd)[1].strip()
            outputs.append(output)
            continue
        output = output.decode('utf-8').split('root@S8P54040023')[0]
        output = output.split(cmd)[1].strip()
        outputs.append(output)
    client.close()
    return outputs


def get_file(remote_path, local_path, filename):
    client = paramiko.Transport(modbus_config['ssh']['ip'], modbus_config['ssh']['port'])
    client.connect(username=modbus_config['ssh']['username'], password=modbus_config['ssh']['password'])
    sftp = paramiko.SFTPClient.from_transport(client)
    src = remote_path + filename
    des = local_path + filename
    sftp.get(src, des)
    client.close()


def put_file(remote_path, local_path, filename):
    client = paramiko.Transport(modbus_config['ssh']['ip'], modbus_config['ssh']['port'])
    client.connect(username=modbus_config['ssh']['username'], password=modbus_config['ssh']['password'])
    sftp = paramiko.SFTPClient.from_transport(client)
    src = remote_path + filename
    des = local_path + filename
    sftp.put(des, src)
    client.close()
