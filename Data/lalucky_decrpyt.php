<?php
$this->load->library('encrypt');

$encrypted_password = 'c56fac72273c88a8d0bcea591bb50b12';
$key = 'lalucky';

$decrypted_string = $this->encrypt->decode($encrypted_password, $key);

echo $decrypted_string;

?>
