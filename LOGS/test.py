import logging

logging.basicConfig(
    filename="LOGS/file.log",
    filemode="a",
    format="%(asctime)s: %(msecs)d %(levelname)s : %(message)s",
    datefmt="%H:%M:%S",
    level=logging.DEBUG,
)

logger = logging.getLogger("my_logger")

def main():
    logging.info("pass the test")

main()
