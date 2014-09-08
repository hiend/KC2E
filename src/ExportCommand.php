<?php
namespace KC2E;

use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;

class ExportCommand extends Command
{
    protected function configure()
    {
        $this
            ->setName('export')
            ->setDescription('Export keywords groups from KeyCollector project')
            ->addArgument(
                'kcdb',
                InputArgument::REQUIRED,
                'KeyCollector file'
            )
            ->addArgument(
                'xls',
                InputArgument::REQUIRED,
                'Excel file'
            )
        ;
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {
        (new Export($input->getArgument('kcdb')))
            ->save($input->getArgument('xls'));

        $output->writeln("Done!");
    }
}
